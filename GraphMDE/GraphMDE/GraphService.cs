using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace GraphMDE
{

    public class GraphService
    {
        private GraphServiceClient? graphServiceClient;
        private string? clientId;
        private string[]? scopes;
        public GraphService(string clientId, string tenantId, string[] scopes)
        {
            this.clientId = clientId;
            this.scopes = scopes;

            InteractiveBrowserCredential credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                ClientId = this.clientId,
                TenantId = tenantId
            });

            this.graphServiceClient = new GraphServiceClient(credential, this.scopes);
        }

        public async Task AddEvent(string subject, string emoji, DateTimeOffset expires)
        {
            var reminderEvent = new Event
            {
                Subject = String.IsNullOrEmpty(emoji) ? $"{subject}" : $"{emoji} {subject}",
                Start = new DateTimeTimeZone
                {
                    DateTime = expires.ToUniversalTime().ToString("o"),
                    TimeZone = TimeZoneInfo.Utc.Id
                },
                End = new DateTimeTimeZone
                {
                    DateTime = expires.ToUniversalTime().AddHours(1).ToString("o"),
                    TimeZone = TimeZoneInfo.Utc.Id
                }
            };

            // Add the event to the user's calendar
            await this.graphServiceClient!.Me.Calendar.Events.PostAsync(reminderEvent);
        }

        public async Task<bool> HasEvent(string subject, DateTimeOffset expires)
        {
            // check if the reminder already exists
            var events = await this.graphServiceClient!.Me.Calendar.Events.GetAsync(r =>
            {
                r.QueryParameters.Filter = $"contains(subject, '{subject}') and start/dateTime eq '{expires.ToUniversalTime().ToString("o")}'";
            });

            return events?.Value?.Count > 0;
        }

        // get user information from microsoft graph
        public async Task<User?> GetMe()
        {
            return await this.graphServiceClient!.Me.GetAsync();
        }

    }
}
