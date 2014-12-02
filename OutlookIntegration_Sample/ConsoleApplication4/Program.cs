using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
namespace ConsoleApplication4
{
    class Program
    {
        static void ListTopLevelFolders(ExchangeService service, int moffset)
        {
            int mPageSize = 5;
            // Create a view with a page size of 5.
            FolderView view = new FolderView(mPageSize, moffset, OffsetBasePoint.Beginning);

            // Identify the properties to return in the results set.
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.PropertySet.Add(FolderSchema.ChildFolderCount);

            // Unlike FindItem searches, folder searches can be deep traversals.
            view.Traversal = FolderTraversal.Shallow;

            // Send the request to search the mailbox and get the results.
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);
            //FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.Calendar, view);

            // Process each item.
            foreach (Folder myFolder in findFolderResults.Folders)
            {
                if (myFolder is SearchFolder)
                {
                    Console.WriteLine("Search folder: " + (myFolder as SearchFolder).DisplayName);
                }

                else if (myFolder is ContactsFolder)
                {
                    Console.WriteLine("Contacts folder: " + (myFolder as ContactsFolder).DisplayName);
                }

                else if (myFolder is TasksFolder)
                {
                    Console.WriteLine("Tasks folder: " + (myFolder as TasksFolder).DisplayName);
                }

                else if (myFolder is CalendarFolder)
                {
                    Console.WriteLine("Calendar folder: " + (myFolder as CalendarFolder).DisplayName);
                    
                }
                else
                {
                    // Handle a generic folder.
                    Console.WriteLine("Folder: " + myFolder.DisplayName);
                }
            }

            // Determine whether there are more folders to return.
            if (findFolderResults.MoreAvailable)
            {
                // Make recursive calls with offsets set for the FolderView to get the remaining folders in the result set.
                moffset = moffset + mPageSize;
                ListTopLevelFolders(service, moffset);
            }
            else
            {
                Console.WriteLine("More Available: " + findFolderResults.MoreAvailable);
            }
        }

        private static void ReadAppoitmentsInCalendar(ExchangeService service)
        {
            DateTime startDate = DateTime.Now;
            DateTime endDate = startDate.AddDays(30);
            const int NUM_APPTS =10;

            
            // Initialize the calendar folder object with only the folder ID. 
            //CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
            //Folder folder = new Folder(service);
            // Folder id for AcutechCalendar
            FolderId folderID = new FolderId("AAMkADg4NTFkZTgwLTE3MGUtNGM0MC1hMjg0LWRlOTE1Yjg5YjMwOQAuAAAAAAD1wws5HI4XRpYtTLjy/6g+AQBUosiWYeNnRbdr4Mf1lS0DABMSUwA6AAA=");
            CalendarFolder calendar = CalendarFolder.Bind(service, folderID, new PropertySet());

            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS);

            // Limit the properties returned to the appointment's subject, start time, and end time.
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.ICalUid);

            // Retrieve a collection of appointments by using the calendar view.
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            Console.WriteLine("\nThe first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
                              " to " + endDate.Date.ToShortDateString() + " are: \n");

            foreach (Appointment a in appointments)
            {
                Console.Write("Subject: " + a.Subject.ToString() + " ");
                Console.Write("Start: " + a.Start.ToString() + " ");
                Console.Write("End: " + a.End.ToString());
                //Console.Write("Calendar id: " + a.ICalUid.ToString());
                Console.WriteLine();
            }
        }
        private static void SendMeetingInvites(ExchangeService service)
        {
            Appointment appointment = new Appointment(service);
            appointment.Subject = "Status Meeting";
            appointment.Body = "The purpose of this meeting is to discuss status.";
            appointment.Start = DateTime.Now;
            appointment.End = appointment.Start.AddHours(2);
            appointment.Location = "Conf Room";
            appointment.RequiredAttendees.Add("raju_bodkhe@persistent.co.in");
            appointment.RequiredAttendees.Add("sangameshwar_todkari@persistent.co.in");
            appointment.OptionalAttendees.Add("sangameshwar_todkari@persistent.co.in");
            appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
            Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));
            Console.WriteLine("\nMeeting created: " + item.Subject + "\n");


        }
        private static void TrackMeetingStatus(ExchangeService service)
        {
            //generated while creating meeting invite . need to captured and stored for future use.
            string itemID = "AAMkADg4NTFkZTgwLTE3MGUtNGM0MC1hMjg0LWRlOTE1Yjg5YjMwOQBGAAAAAAD1wws5HI4XRpYtTLjy/6g+BwCy77FnxTbXQLWhn0pYOwLHAAAAoSIrAACy77FnxTbXQLWhn0pYOwLHAABHaX2HAAA=";
            Appointment meeting = Appointment.Bind(service, new ItemId(itemID));


            // Check responses from required attendees.
            for (int i = 0; i < meeting.RequiredAttendees.Count; i++)
            {
                Console.WriteLine("Required attendee - " + meeting.RequiredAttendees[i].Address + ": " + meeting.RequiredAttendees[i].ResponseType.Value.ToString());
            }

            // Check responses from optional attendees.
            for (int i = 0; i < meeting.OptionalAttendees.Count; i++)
            {
                Console.WriteLine("Optional attendee - " + meeting.OptionalAttendees[i].Address + ": " + meeting.OptionalAttendees[i].ResponseType.Value.ToString());
            }

            // Check responses from resources.
            for (int i = 0; i < meeting.Resources.Count; i++)
            {
                Console.WriteLine("Resource attendee - " + meeting.Resources[i].Address + ": " + meeting.Resources[i].ResponseType.Value.ToString());
            }
        }
        static void Main(string[] args)
        {
            // Create the binding.
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
            // Set the credentials for the on-premises server.
            service.Credentials = new WebCredentials("sangameshwar_todkari@persistent.co.in", "password");

            // Set the URL.
            service.Url = new Uri("https://mail.persistent.co.in/EWS/Exchange.asmx");

            //ListTopLevelFolders(service, 0);
            
            ReadAppoitmentsInCalendar(service);
            //SendMeetingInvites(service);
            //TrackMeetingStatus(service);
        }

    }
}
