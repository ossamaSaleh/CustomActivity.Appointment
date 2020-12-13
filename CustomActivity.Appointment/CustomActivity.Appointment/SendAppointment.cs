using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using System.Net.Mail;

namespace CustomActivity.Appointment
{
    public class SendAppointment : CodeActivity
    {
        [RequiredArgument]
        [Input("Start Date")]
        public InArgument<DateTime> From { get; set; }
        [RequiredArgument]
        [Input("EndDate")]
        public InArgument<DateTime> To { get; set; }
        [RequiredArgument]
        [Input("Email")]
        public InArgument<string> Email { get; set; }
        [RequiredArgument]
        [Input("Subject")]
        public InArgument<string> Subject { get; set; }
        [RequiredArgument]
        [Input("Location")]
        public InArgument<string> Location { get; set; }
        [RequiredArgument]
        [Input("EnableSSL")]
        public InArgument<bool> IsEnableSSL { get; set; }
        [RequiredArgument]
        [Input("SMTP Configuration")]
        [ReferenceTarget("os_emailconfiguration")]
        public InArgument<EntityReference> Configuration { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            ITracingService tracingService = context.GetExtension<ITracingService>();
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            var orgService = serviceFactory.CreateOrganizationService(workflowContext.UserId);
            DateTime AcStartDate = this.From.Get<DateTime>(context);
            DateTime AcEndDate = this.To.Get<DateTime>(context);
            string _email = this.Email.Get<string>(context);
            string _subject = this.Subject.Get<string>(context);
            string _location = this.Location.Get<string>(context);
            var _isSSL = this.IsEnableSSL.Get<bool>(context);
            //Get the Configuration
            EntityReference _config = this.Configuration.Get<EntityReference>(context);
            var ConfigData = orgService.Retrieve(_config.LogicalName, _config.Id, new ColumnSet(true));
            var _smtpIP = ConfigData.GetAttributeValue<string>("os_smtpip");
            var _username = ConfigData.GetAttributeValue<string>("os_username");
            var _pass = ConfigData.GetAttributeValue<string>("os_password");
            var _sender = ConfigData.GetAttributeValue<string>("emailaddress");
            var _orgName = ConfigData.GetAttributeValue<string>("os_name");
            var _port = ConfigData.GetAttributeValue<int>("os_port");
            IICalendar iCal = new iCalendar();
            iCal.Method = "Request";
            IEvent evt = iCal.Create<Event>();
            evt.Summary = _subject;
            evt.Start = new iCalDateTime(AcStartDate);
            evt.End = new iCalDateTime(AcEndDate);
            evt.Description = _subject;
            evt.Location = _location;
            tracingService.Trace("evt done");
            //evt.Name = _subject;
            //evt.Organizer = new Organizer("MAILTO:" + _sender);
            //evt.Organizer.CommonName = _orgName;
            Alarm alarm = new Alarm();

            // Display the alarm somewhere on the screen.
            alarm.Action = AlarmAction.Display;

            // This is the text that will be displayed for the alarm.
            alarm.Summary = _subject;

            // The alarm is set to occur 30 minutes before the event
            alarm.Trigger = new Trigger(TimeSpan.FromMinutes(-30));
            var serializer = new iCalendarSerializer(iCal);
            var iCalString = serializer.SerializeToString(iCal);
            var msg = new MailMessage();
            msg.Subject = _subject;
            msg.Body = _subject;
            msg.From = new MailAddress(_sender);
            msg.To.Add(_email);

            tracingService.Trace("mail done");
            // Create the Alternate view object with Calendar MIME type
            var ct = new System.Net.Mime.ContentType("text/calendar");
            if (ct.Parameters != null) ct.Parameters.Add("method", "REQUEST");

            //Provide the framed string here
            AlternateView avCal = AlternateView.CreateAlternateViewFromString(iCalString, ct);
            msg.AlternateViews.Add(avCal);

            // Send email
            try
            {
                SendSMTP(_smtpIP, _port, _username, _pass, _isSSL);
                tracingService.Trace("done");

            }
            catch (Exception ex) { throw new InvalidPluginExecutionException("Error " + ex.Message); }

        }
        void SendSMTP(string SMTPHost, int Port, string Username, string Password, bool isSSL)
        {
            SmtpClient smtp = new SmtpClient();
            smtp.Host = SMTPHost;

            smtp.Port = Port;
            smtp.UseDefaultCredentials = false;
            //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

            if (isSSL)
                smtp.EnableSsl = true;
            smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            smtp.Credentials = new System.Net.NetworkCredential(Username, Password);
        }
    }
}
