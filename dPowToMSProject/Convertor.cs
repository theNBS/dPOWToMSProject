using System;
using System.Collections.Generic;
using System.Linq;

// Xbim DPoW library contains functionality for parsing and interacting with dPOW files
using Xbim.DPoW;

// MPXJ library allows reading and writing of MS Project formats (including XML)
// This library uses IVKM (.NET implementation of Java), since the original library was written in Java
using net.sf.mpxj.writer;
using net.sf.mpxj;
using java.util;
using java.text;
using java.lang;

namespace dPowToMSProject
{
    /// <summary>
    /// Provides the ability to convert a dPOW file into a MS Project XML file to allow viewing of tasks and resources or a Gantt chart
    /// </summary>
    public class Convertor
    {
        // Name of calendar to use as default calendar for the project
        private const string PROJECT_CALENDER_NAME = "24 Hour Calendar";

        // Path to the dPOW file
        private string inputFilePath;

        // Store added resources (i.e. contacts) in a dictionary to avoid adding duplicates
        private Dictionary<Guid, Resource> addedResources = new Dictionary<Guid, Resource>();

        // DPOW object
        PlanOfWork pow;

        // Project object
        ProjectFile project;

        /// <summary>
        /// Instantiate the convertor object
        /// </summary>
        /// <param name="inputFilePath">Path of dPOW file to read</param>
        public Convertor(string inputFilePath)
        {
            this.inputFilePath = inputFilePath;
        }

        /// <summary>
        /// Perform the conversion, writing data to the provided output path
        /// </summary>
        /// <param name="outputFilePath">Path of output xml file (should have the extension .xml so the MPXJ library automatically infers the file type)</param>
        public void Convert(string outputFilePath)
        {
            // Open and parse dPOW file
            pow = PlanOfWork.OpenJson(inputFilePath);

            // Create project file
            project = new ProjectFile();

            // Insert data to project file
            InsertDataToProject();

            // Write file to mpx
            ProjectWriter mpxWriter = ProjectWriterUtility.getProjectWriter(outputFilePath);
            mpxWriter.write(project, outputFilePath);
        }

        /// <summary>
        /// Insert data from the dPOW to the new project file
        /// </summary>
        private void InsertDataToProject()
        {
            // Set main project settings (name, title etc.)
            project.ProjectProperties.Name = pow.Project.Name;
            project.ProjectProperties.ProjectTitle = pow.Project.Name;
            project.ProjectProperties.Comments = pow.Project.Description;

            // dPOW does not take into account work hours or days, so set calendar and scheduling settings for 24 hours (otherwise calculated durations and end dates will not be as expected) 
            SetupCalendar(project.AddDefaultBaseCalendar());
            project.ProjectProperties.DefaultCalendarName = PROJECT_CALENDER_NAME;

            // In addition to setting the calendar, the project scheduling settings must be updated for 24 hours
            project.ProjectProperties.DefaultStartTime = new Date(0); // Start time to 00:00:00
            project.ProjectProperties.DefaultEndTime = new Date(0); // End time to 00:00:00
            project.ProjectProperties.MinutesPerDay = new Integer(1440); // 24 hours per day
            project.ProjectProperties.MinutesPerWeek = new Integer(10080); // 7 days per week

            // Add each dPOW project stages as milestone tasks
            foreach(var stage in pow.ProjectStages)
            {
                // Create task from project stage
                // These are the top level summary tasks, which all other tasks are children of
                Task task = project.AddTask();
                PopulateTask(task, stage.Name, stage.Description, stage.Attributes);
                task.Summary = true;
                task.Milestone = true;

                // Add dPOW document sets to the project stage tasks as subtasks
                foreach(var documentSet in stage.DocumentationSet)
                {
                    // Create task from documentation set
                    Task documentSetTask = task.AddTask();
                    PopulateTask(documentSetTask, documentSet.Name, documentSet.Description, documentSet.Attributes, task);

                    // Iterate jobs to find (currently there is only one job per document set, which contains the responsible contact and role)
                    foreach(var job in documentSet.Jobs)
                    {
                        SetupResourceAssignment(documentSetTask, job);
                    }
                }

                // Add other dPOW task types (assets, assembly and space deliverables) to the project stage task
                foreach(var asset in stage.AssetTypes)
                {
                    PopulateTask(task.AddTask(), asset.Name, asset.Description, asset.Attributes, task);
                }
                foreach(var assembly in stage.AssemblyTypes)
                {
                    PopulateTask(task.AddTask(), assembly.Name, assembly.Description, assembly.Attributes, task);
                }
                foreach(var space in stage.SpaceTypes)
                {
                    PopulateTask(task.AddTask(), space.Name, space.Description, space.Attributes, task);
                }
            }
        }

        #region helper methods for retreiving and adapting data

        /// <summary>
        /// Creates the resource assigmnet for the task, adding the resource to the project if it has not been added yet
        /// </summary>
        /// <param name="task">Task to assign resource to</param>
        /// <param name="resource">Resource to assign to task</param>
        private void SetupResourceAssignment(Task task, Job job)
        {
            // Add a resource assignment to the task for responsible contacts
            Guid contactGuid = job.Responsibility.ResponsibleContactId;
            if (addedResources.ContainsKey(contactGuid))
            {
                // Resource has already been added to project so assign existing resource to task
                task.AddResourceAssignment(addedResources[contactGuid]);
            }
            else
            {
                // Resource has not been added - retrieve contact details, add resource to project and assign the new resource to the task
                Contact contact = pow.Contacts.FirstOrDefault(c => c.Id == contactGuid);
                if (contact != null)
                {
                    Resource resource = project.AddResource();
                    PopulateResourceFromContact(resource, contact);
                    addedResources.Add(contactGuid, resource);
                    task.AddResourceAssignment(resource);
                }
            }
        }

        /// <summary>
        /// Populate a provided Resource project 
        /// </summary>
        /// <param name="resource">Resource object to populate with contact details</param>
        /// <param name="contact">dPOW contact object containing the details</param>
        private static void PopulateResourceFromContact(Resource resource, Contact contact)
        {
            // Assign resource contact name and email address
            // If contact does not have a GivenName (appears to often be the case), fall back on the CompanyName
            // MS Project reports errors if the name contains a '[' or ']', which are found in the '[Not Assigned]' contact, so replace these with normal brackets
            resource.Name = (contact.GivenName ?? contact.CompanyName).Replace("[", "(").Replace("]", ")");
            resource.EmailAddress = contact.Email;
        }

        /// <summary>
        /// Populate a Task object with the provided values
        /// </summary>
        /// <param name="task">task object for populating</param>
        /// <param name="name">name of the task</param>
        /// <param name="notes">notes to attach to the task</param>
        /// <param name="attributes">list of dPOW attributes to get values from</param>
        /// <param name="parent">(optional) pass parent to allow inheritance of start and end dates and deadlines</param>
        private static void PopulateTask(Task task, string name, string notes, List<Xbim.DPoW.Attribute> attributes, Task parent = null)
        {
            task.Name = name;
            // Some dPOW objects description value, others have notes in the attributes list 
            task.Notes = notes ?? StringFromAttribute(attributes, "Notes");
            task.Cost = NumberFromAttribute(attributes, "Cost", new Integer(0));
            // Try to get dates from attributes, if none available, and parent task exists, inherit from parent 
            task.Start = DateFromAttributes(attributes, "StartDate", parent != null ? parent.Start : null);
            task.Finish = DateFromAttributes(attributes, "EndDate", parent != null ? parent.Finish : null);
            task.Deadline = DateFromAttributes(attributes, "StageDeadline", null);
            // Calculate and set project duration from the start and finish date
            Duration duration = DurationBetween(task.Start, task.Finish);
            task.ManualDuration = duration;
            task.Duration = duration;
            // The following task options should be set to ensure it is scheduled correctly
            task.TaskMode = TaskMode.MANUALLY_SCHEDULED;
            task.IgnoreResourceCalendar = true;
        }

        /// <summary>
        /// Takes a list of attributes from a dPOW object and tries to parse the value of the given key into a date
        /// </summary>
        /// <param name="attributesList">List of attributes in dPOW object</param>
        /// <param name="key">Key to return the parsed value of</param>
        /// <param name="defaultValue">Value to return if date cannot be found or parsed</param>
        /// <returns>Date object or defaultValue if date could not be found or parsed</returns>
        private static java.util.Date DateFromAttributes(List<Xbim.DPoW.Attribute> attributesList, string key, Date defaultValue)
        {
            // Search the attributes list for an attribute matching the given key
            if(attributesList.Any(a => a.Name == key))
            {
                // Attempt to parse a date from the value
                DateFormat format = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
                try
                {
                    return format.parse(attributesList.First(a => a.Name == key).Value);
                }
                catch(System.Exception e)
                {
                    return defaultValue;
                }
            }

            return defaultValue;
        }

        /// <summary>
        /// Takes a list of attributes from a dPOW object and tries to parse the value of the given key into a number
        /// </summary>
        /// <param name="attributesList">List of attributes in dPOW object</param>
        /// <param name="key">Key to return the parsed value of</param>
        /// <param name="defaultValue">Value to return if number cannot be found or parsed</param>
        /// <returns>Number object or null if date could not be found or parsed</returns>
        private static Number NumberFromAttribute(List<Xbim.DPoW.Attribute> attributesList, string key, Number defaultValue)
        {
            // Search the attributes list for an attribute matching the given key
            if (attributesList.Any(a => a.Name == key))
            {
                // Attempt to parse a number from the value
                NumberFormat format = NumberFormat.getNumberInstance(Locale.ENGLISH);
                try {
                    return format.parse(attributesList.First(a => a.Name == key).Value);
                }
                catch(System.Exception e)
                {
                    return defaultValue;
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="attributesList">List of attributes in dPOW object</param>
        /// <param name="key">Key to return the parsed value of</param>
        /// <returns>Value of attribute or null if attribute not found</returns>
        private static string StringFromAttribute(List<Xbim.DPoW.Attribute> attributesList, string key)
        {
            // If the attribute list contains an attribute matching the given key, return the associated value, otherwise return null
            return attributesList.Any(a => a.Name == key) ? attributesList.First(a => a.Name == key).Value : null;
        }

        /// <summary>
        /// Return the duration between two dates. If either date is null a duration of 0 is returned. 
        /// </summary>
        /// <param name="start">start date</param>
        /// <param name="finish">end date</param>
        /// <returns>duration between the dates (or duration of 0 if either date is null)</returns>
        private static Duration DurationBetween(Date start, Date finish)
        {
            double duration = 0;
            if(start != null && finish != null)
            {
                duration = finish.getTime() - start.getTime();
            }
            return Duration.getInstance(duration / 1000 / 60 / 60 / 24, TimeUnit.DAYS);
        }

        #endregion

        #region calendar setup

        /// <summary>
        /// Set up a 24 hour calendar
        /// </summary>
        /// <param name="calendar"></param>
        private void SetupCalendar(ProjectCalendar calendar)
        {
            // Simple date format for setting dates
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm");

            // Date range containing all hours
            DateRange range = new DateRange(format.parse("2000-01-01 00:00"), format.parse("2000-01-02 00:00"));

            // Set calendar name to the same value to which the project calendar will be set
            calendar.Name = PROJECT_CALENDER_NAME;

            // Mark each day as working
            calendar.setWorkingDay(Day.SUNDAY, true);
            calendar.setWorkingDay(Day.MONDAY, true);
            calendar.setWorkingDay(Day.TUESDAY, true);
            calendar.setWorkingDay(Day.WEDNESDAY, true);
            calendar.setWorkingDay(Day.THURSDAY, true);
            calendar.setWorkingDay(Day.FRIDAY, true);
            calendar.setWorkingDay(Day.SATURDAY, true);

            // Add a working hours range to each day
            ProjectCalendarHours hours;
            hours = calendar.addCalendarHours(Day.SUNDAY);
            hours.addRange(range);
            hours = calendar.addCalendarHours(Day.MONDAY);
            hours.addRange(range);
            hours = calendar.addCalendarHours(Day.TUESDAY);
            hours.addRange(range);
            hours = calendar.addCalendarHours(Day.WEDNESDAY);
            hours.addRange(range);
            hours = calendar.addCalendarHours(Day.THURSDAY);
            hours.addRange(range);
            hours = calendar.addCalendarHours(Day.FRIDAY);
            hours.addRange(range);
            hours = calendar.addCalendarHours(Day.SATURDAY);
            hours.addRange(range);
        }

        #endregion
    }
}
