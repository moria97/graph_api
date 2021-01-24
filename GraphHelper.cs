using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace graph_api
{
    public class GraphHelper
    {
        private static GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

        public static async Task<User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await graphClient.Me
                    .Request()
                    .Select(u => new{
                        u.DisplayName,
                        u.MailboxSettings
                    })
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        public static async Task SendMailAsync(string toUser, string title, string content)
        {
            Message message = new Message
            {
                Subject = title,
                Body = new ItemBody { Content = content },
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = toUser
                        }
                    }
                }
            };

            try
            {
                await graphClient.Me.SendMail(message).Request().PostAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error sending mail: {ex.Message}");
            }
        }
    }
}