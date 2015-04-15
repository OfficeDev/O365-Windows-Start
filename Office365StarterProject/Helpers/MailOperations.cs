using Microsoft.Office365.OutlookServices;
using Office365StarterProject.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365StarterProject.Helpers
{
    class MailOperations
    {
        private string _mailCapability = ServiceCapabilities.Mail.ToString();

        /// <summary>
        /// Fetches email from user's Inbox.
        /// </summary>
        /// <param name="pageNo">The page of mail results to be fetched.</param>
        /// <param name="bodyContent">The size of the results page.</param>
        internal async Task<List<Message>> GetEmailMessagesAsync(int pageNo, int pageSize)
        {

            // Make sure we have a reference to the Outlook Services client
            var outlookClient = await AuthenticationHelper.GetOutlookClientAsync(_mailCapability);

            var mailResults = await (from i in outlookClient.Me.Folders.GetById("Inbox").Messages
                                     orderby i.DateTimeReceived descending
                                     select i).Skip((pageNo - 1) * pageSize).Take(pageSize).ExecuteAsync();

            foreach (var message in mailResults.CurrentPage)
            {
                  System.Diagnostics.Debug.WriteLine("Message '{0}' received at '{1}'.",
                  message.Subject,
                  message.DateTimeReceived.ToString());
            }


            return (List<Message>)mailResults.CurrentPage;

        }

        /// <summary>
        /// Compose and send a new email.
        /// </summary>
        /// <param name="subject">The subject line of the email.</param>
        /// <param name="bodyContent">The body of the email.</param>
        /// <param name="recipients">A semicolon separated list of email addresses.</param>
        /// <returns></returns>
        internal async Task<String> ComposeAndSendMailAsync(string subject,
                                                            string bodyContent,
                                                            string recipients)
        {
            // The identifier of the composed and sent message.
            string newMessageId = string.Empty;

            // Prepare the recipient list
            var toRecipients = new List<Recipient>();
            string[] splitter = { ";" };
            var splitRecipientsString = recipients.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
            foreach (string recipient in splitRecipientsString)
            {
                toRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient.Trim(),
                        Name = recipient.Trim(),
                    },
                });
            }

            // Prepare the draft message.
            var draft = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = bodyContent
                },
                ToRecipients = toRecipients,
            };

            try
            {
                // Make sure we have a reference to the Outlook Services client.
                var outlookClient = await AuthenticationHelper.GetOutlookClientAsync(_mailCapability);

                //Send the mail.
                await outlookClient.Me.SendMailAsync(draft, true);

                return draft.Id;
            }

            //Catch any exceptions related to invalid OData.
            catch (Microsoft.OData.Core.ODataException ode)
            {

                throw new Exception("We could not send the message: " + ode.Message);
            }
            catch (Exception e)
            {
                throw new Exception("We could not send the message: " + e.Message);
            }
        }

        /// <summary>
        /// Removes a mail item from the user's inbox.
        /// </summary>
        /// <param name="selectedMailId">string. The unique Id of the mail item to delete.</param>
        /// <returns></returns>
        internal async Task<bool> DeleteMailItemAsync(string selectedMailId)
        {
            IMessage thisMailItem = null;
            try
            {
                // Make sure we have a reference to the Outlook Services client
                var outlookClient = await AuthenticationHelper.GetOutlookClientAsync(_mailCapability);

                // Get the mail item to be removed.
                thisMailItem = await outlookClient.Me.Folders.GetById("Inbox").Messages.GetById(selectedMailId).ExecuteAsync();

                // Delete the mail item.
                await thisMailItem.DeleteAsync(false);
                return true;
            }

            //Catch any exceptions related to invalid OData.
            catch (Microsoft.OData.Core.ODataException ode)
            {

                throw new Exception("The message could not be deleted: " + ode.Message);
            }

            catch (Exception e)
            {
                throw new Exception("The message could not be deleted: " + e.Message);
            }

        }

        internal string BuildRecipientList(IList<Recipient> recipientList)
        {
            StringBuilder recipientListBuilder = new StringBuilder();
            foreach (Recipient recipient in recipientList)
            {
                if (recipientListBuilder.Length == 0)
                {
                    recipientListBuilder.Append(recipient.EmailAddress.Address);
                }
                else
                {
                    recipientListBuilder.Append(";" + recipient.EmailAddress.Address);
                }
            }

            return recipientListBuilder.ToString();
        }

    }
}

//********************************************************* 
// 
//O365-APIs-Start-Windows, https://github.com/OfficeDev/O365-APIs-Start-Windows
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 