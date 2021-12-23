// <copyright file="SendFunction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;
    using Newtonsoft.Json;
    using Attachment = Microsoft.Bot.Schema.Attachment;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue
    /// Used for sending messages from the bot.
    /// </summary>
    public class SendFunction
    {
        /// <summary>
        /// This is set to 10 because the default maximum delivery count from the service bus
        /// message queue before the service bus will automatically put the message in the Dead Letter
        /// Queue is 10.
        /// </summary>
        private static readonly int MaxDeliveryCountForDeadLetter = 10;
        private static readonly string AdaptiveCardContentType = "application/vnd.microsoft.card.adaptive";

        private readonly string emailSenderAadId;
        private readonly string tenantId;
        private readonly string originatorId;
        private readonly string authorAppId;
        private readonly string authorAppPassword;
        private readonly string appServiceUri;

        private readonly int maxNumberOfAttempts;
        private readonly double sendRetryDelayNumberOfSeconds;
        private readonly INotificationService notificationService;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISendingNotificationDataRepository notificationRepo;
        private readonly IMessageService messageService;
        private readonly ISendQueue sendQueue;
        private readonly IStringLocalizer<Strings> localizer;

        // Test Email Option Start
        string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
        AdaptiveCard aCard;
        AdaptiveCard teamsCard;

        // Test Email Option End

        /// <summary>
        /// Initializes a new instance of the <see cref="SendFunction"/> class.
        /// </summary>
        /// <param name="options">Send function options.</param>
        /// <param name="notificationService">The service to precheck and determine if the queue message should be processed.</param>
        /// <param name="notificationDataRepository">The service to precheck and determine should be send email.</param>
        /// <param name="messageService">Message service.</param>
        /// <param name="notificationRepo">Notification repository.</param>
        /// <param name="sendQueue">The send queue.</param>
        /// <param name="localizer">Localization service.</param>
        public SendFunction(
            IOptions<SendFunctionOptions> options,
            INotificationService notificationService,
            INotificationDataRepository notificationDataRepository, // Testing Check Email Option
            IMessageService messageService,
            ISendingNotificationDataRepository notificationRepo,
            ISendQueue sendQueue,
            IStringLocalizer<Strings> localizer)
        {
            if (options is null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            this.emailSenderAadId = options.Value.EmailSenderAadId;
            this.tenantId = options.Value.TenantId;
            this.originatorId = options.Value.OriginatorId;
            this.authorAppId = options.Value.AuthorAppId;
            this.authorAppPassword = options.Value.AuthorAppPassword;
            this.appServiceUri = options.Value.AppServiceUri;

            this.maxNumberOfAttempts = options.Value.MaxNumberOfAttempts;
            this.sendRetryDelayNumberOfSeconds = options.Value.SendRetryDelayNumberOfSeconds;

            this.notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentException(nameof(notificationDataRepository)); // Testing Check Email Option
            this.messageService = messageService ?? throw new ArgumentNullException(nameof(messageService));
            this.notificationRepo = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.sendQueue = sendQueue ?? throw new ArgumentNullException(nameof(sendQueue));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
        }

        /// <summary>
        /// Azure Function App triggered by messages from a Service Bus queue
        /// Used for sending messages from the bot.
        /// </summary>
        /// <param name="myQueueItem">The Service Bus queue item.</param>
        /// <param name="deliveryCount">The deliver count.</param>
        /// <param name="enqueuedTimeUtc">The enqueued time.</param>
        /// <param name="messageId">The message ID.</param>
        /// <param name="log">The logger.</param>
        /// <param name="context">The execution context.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName("SendMessageFunction")]
        public async Task Run(
            [ServiceBusTrigger(
                SendQueue.QueueName,
                Connection = SendQueue.ServiceBusConnectionConfigurationKey)]
            string myQueueItem,
            int deliveryCount,
            DateTime enqueuedTimeUtc,
            string messageId,
            ILogger log,
            ExecutionContext context)
        {
            log.LogInformation($"C# ServiceBus queue trigger function processed message: {myQueueItem}");

            var messageContent = JsonConvert.DeserializeObject<SendQueueMessageContent>(myQueueItem);

            try
            {
                // Init Graph.
                IConfidentialClientApplication confidentialClient = ConfidentialClientApplicationBuilder
                .Create(this.authorAppId)
                .WithClientSecret(this.authorAppPassword)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/" + this.tenantId + "/v2.0"))
                .Build();

                // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                var authResult = await confidentialClient
                        .AcquireTokenForClient(this.scopes)
                        .ExecuteAsync().ConfigureAwait(false);

                var token = authResult.AccessToken;

                // Build the Microsoft Graph client. As the authentication provider, set an async lambda
                // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
                // and inserts this access token in the Authorization header of each API request. 
                GraphServiceClient graphServiceClient =
                    new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization =
                                    new AuthenticationHeaderValue("Bearer", token);
                    }));

                // Check if recipient is a guest user.
                if (messageContent.IsRecipientGuestUser())
                {
                    await this.notificationService.UpdateSentNotification(
                        notificationId: messageContent.NotificationId,
                        recipientId: messageContent.RecipientData.RecipientId,
                        totalNumberOfSendThrottles: 0,
                        statusCode: SentNotificationDataEntity.NotSupportedStatusCode,
                        allSendStatusCodes: $"{SentNotificationDataEntity.NotSupportedStatusCode},",
                        errorMessage: this.localizer.GetString("GuestUserNotSupported"),
                        reactionResult: "",
                        freeTextResult: "",
                        yesNoResult: "");
                    return;
                }

                // Check if notification is pending.
                var isPending = await this.notificationService.IsPendingNotification(messageContent);
                if (!isPending)
                {
                    // Notification is either already sent or failed and shouldn't be retried.
                    return;
                }

                // Check if conversationId is set to send message.
                if (string.IsNullOrWhiteSpace(messageContent.GetConversationId()))
                {
                    await this.notificationService.UpdateSentNotification(
                        notificationId: messageContent.NotificationId,
                        recipientId: messageContent.RecipientData.RecipientId,
                        totalNumberOfSendThrottles: 0,
                        statusCode: SentNotificationDataEntity.FinalFaultedStatusCode,
                        allSendStatusCodes: $"{SentNotificationDataEntity.FinalFaultedStatusCode},",
                        errorMessage: this.localizer.GetString("AppNotInstalled"),
                        reactionResult: "",
                        freeTextResult: "",
                        yesNoResult: "");
                    return;
                }

                // Check if the system is throttled.
                var isThrottled = await this.notificationService.IsSendNotificationThrottled();
                if (isThrottled)
                {
                    // Re-Queue with delay.
                    await this.sendQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                    return;
                }

                // Send message.
                var messageActivity = await this.GetMessageActivity(messageContent);
                var response = await this.messageService.SendMessageAsync(
                    message: messageActivity,
                    serviceUrl: messageContent.GetServiceUrl(),
                    conversationId: messageContent.GetConversationId(),
                    maxAttempts: this.maxNumberOfAttempts,
                    logger: log);

                // Process response.
                await this.ProcessResponseAsync(messageContent, response, log);

                // Send Adaptive Card to Email.
                var notificationId = messageContent.NotificationId;
                var notificationEntity = await this.notificationDataRepository.GetAsync(NotificationDataTableNames.SentNotificationsPartition, notificationId); // Testing Check Email Option
                var recData = messageContent.RecipientData.RecipientId;
                if (notificationEntity.EmailOption)
                {
                    //string tJson = "{\"type\":\"AdaptiveCard\",\"originator\":\"ae832c7e-ad1e-4fda-9c9d-e9a98ac84dfb\",\"version\":\"1.0\",\"body\":[{\"type\":\"Container\",\"backgroundImage\":{\"url\":\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC\",\"fillMode\":\"repeat\",\"horizontalAlignment\":\"center\",\"verticalAlignment\":\"center\"},\"items\":[{\"type\":\"TextBlock\",\"size\":\"medium\",\"weight\":\"bolder\",\"color\":\"light\",\"text\":\"Demo\",\"horizontalAlignment\":\"center\"}],\"bleed\":true},{\"type\":\"TextBlock\",\"size\":\"extraLarge\",\"weight\":\"bolder\",\"text\":\"Outlook Client Demo \",\"wrap\":true},{\"type\":\"Image\",\"size\":\"stretch\",\"url\":\"https://i.ytimg.com/vi/rMXAl05wtzQ/maxresdefault.jpg\",\"altText\":\"\"},{\"type\":\"TextBlock\",\"size\":\"medium\",\"weight\":\"bolder\",\"text\":\"Summary Demo\nSummary Demo\nSummary Demo\nSummary Demo\",\"horizontalAlignment\":\"center\",\"wrap\":true,\"fontType\":\"monospace\"},{\"type\":\"TextBlock\",\"size\":\"small\",\"weight\":\"lighter\",\"text\":\"Demo\",\"wrap\":true},{\"type\":\"ActionSet\",\"actions\":[{\"type\":\"Action.OpenUrl\",\"url\":\"https://office.com\",\"title\":\"Read More\"}]},{\"type\":\"TextBlock\",\"text\":\"Demo\",\"wrap\":true},{\"type\":\"Input.ChoiceSet\",\"id\":\"Reaction\",\"style\":\"expanded\",\"isMultiSelect\":false,\"choices\":[{\"title\":\"Extremely satisfied\",\"value\":\"1\"},{\"title\":\"Somewhat satisfied\",\"value\":\"2\"},{\"title\":\"Neither satisfied nor dissatisfied\",\"value\":\"3\"},{\"title\":\"Somewhat dissatisfied\",\"value\":\"4\"},{\"title\":\"Extremely dissatisfied\",\"value\":\"5\"}]},{\"type\":\"TextBlock\",\"text\":\"Demo\",\"wrap\":true},{\"type\":\"Input.Text\",\"id\":\"FreeTextSurvey\",\"placeholder\":\"Enter Text Here\",\"isMultiline\":true,\"maxLength\":500},{\"type\":\"TextBlock\",\"text\":\"Demo\",\"wrap\":true},{\"type\":\"Input.ChoiceSet\",\"id\":\"YesNo\",\"style\":\"expanded\",\"isMultiSelect\":false,\"choices\":[{\"title\":\"Yes\",\"value\":\"Yes\"},{\"title\":\"No\",\"value\":\"No\"}]},{\"type\":\"ActionSet\",\"actions\":[{\"type\":\"Action.Http\",\"method\":\"GET\",\"url\":\"https://chrischow.ap.ngrok.io/api/Survey/Result?notificationId=2517668294588227007&aadid=19baaacc-7c87-47f6-a399-77ceb5d28de1&reaction={{{{Reaction.value}}}}&freetext={{{{FreeTextSurvey.value}}}}&yesno={{{{YesNo.value}}}}\",\"data\":{\"notificationId\":\"2517668294588227007\"},\"title\":\"Submit\"}]},{\"type\":\"ActionSet\",\"actions\":[{\"type\":\"Action.OpenUrl\",\"url\":\"https://office.com\",\"title\":\"Open Survey\"}]}]}";
                    string teamsJson = this.teamsCard.ToJson();
                    string json = this.aCard.ToJson()
                    .Replace("\"type\": \"AdaptiveCard\",", $"\"type\": \"AdaptiveCard\",\"originator\":\"{this.originatorId}\",")
                    .Replace("\"version\": \"1.2\",", $"\"version\": \"1.0\",\"autoInvokeAction\": {{\"method\": \"GET\",\"url\": \"{this.appServiceUri}/api/GetUpdatedCard/Result?notificationId={messageContent.NotificationId}&aadid={messageContent.RecipientData.RecipientId}\",\"body\": \"\",\"type\":\"Action.Http\"}},")
                    .Replace("\\n", "\\n\\r")
                    .Replace("&lt;", "<")
                    .Replace("&gt;", ">")
                    .Replace("&quot;", "&ldquo;")
                    .Replace("&amp;", "&")
                    .Replace("&#39;", "'")
                    .Replace("\"type\": \"Action.Submit\",", $"\"type\": \"Action.Http\",\"method\": \"GET\", \"url\": \"{this.appServiceUri}/api/Survey/Result?notificationId={notificationId}&aadid={recData}&reaction={{{{Reaction.value}}}}&freetext={{{{FreeTextSurvey.value}}}}&yesno={{{{YesNo.value}}}}\",");
                    var sendMail2User = await graphServiceClient.Users[recData]
                        .Request()
                        .Select("userPrincipalName")
                        .GetAsync();
                    var message = new Message
                    {
                        Subject = "Company Communicator: " + notificationEntity.Title,
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><script type='application/adaptivecard+json'>" + json + "</script></head><body>If you are not able to see this mail, click <a href='https://outlook.office.com/mail/inbox'>here</a> to check in Outlook Web Client.<br></body></html>",
                            //Content = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><script type='application/adaptivecard+json'>" + tJson + "</script></head><body>If you are not able to see this mail, click <a href='https://outlook.office.com/mail/inbox'>here</a> to check in Outlook Web Client.<br>" + JsonConvert.SerializeObject(tJson, Formatting.Indented) + "</body></html>",
                        },
                        ToRecipients = new List<Recipient>()
                        {
                            new Recipient
                            {
                                EmailAddress = new EmailAddress
                                {
                                    Address = sendMail2User.UserPrincipalName,
                                },
                            },
                        },
                    };

                    await graphServiceClient.Users[this.emailSenderAadId]
                          .SendMail(message, false)
                          .Request()
                          .PostAsync();
                    return;
                }
            }
            catch (InvalidOperationException exception)
            {
                // Bad message shouldn't be requeued.
                log.LogError(exception, $"InvalidOperationException thrown. Error message: {exception.Message}");
            }
            catch (Exception e)
            {
                var errorMessage = $"{e.GetType()}: {e.Message}";
                log.LogError(e, $"Failed to send message. ErrorMessage: {errorMessage}");

                // Update status code depending on delivery count.
                var statusCode = SentNotificationDataEntity.FaultedAndRetryingStatusCode;
                if (deliveryCount >= SendFunction.MaxDeliveryCountForDeadLetter)
                {
                    // Max deliveries attempted. No further retries.
                    statusCode = SentNotificationDataEntity.FinalFaultedStatusCode;
                }

                // Update sent notification table.
                await this.notificationService.UpdateSentNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: 0,
                    statusCode: statusCode,
                    allSendStatusCodes: $"{statusCode},",
                    errorMessage: errorMessage,
                    reactionResult: "",
                    freeTextResult: "",
                    yesNoResult: "");

                throw;
            }
        }

        /// <summary>
        /// Process send notification response.
        /// </summary>
        /// <param name="messageContent">Message content.</param>
        /// <param name="sendMessageResponse">Send notification response.</param>
        /// <param name="log">Logger.</param>
        private async Task ProcessResponseAsync(
            SendQueueMessageContent messageContent,
            SendMessageResponse sendMessageResponse,
            ILogger log)
        {
            if (sendMessageResponse.ResultType == SendMessageResult.Succeeded)
            {
                log.LogInformation($"Successfully sent the message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}");
            }
            else
            {
                log.LogError($"Failed to send message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {sendMessageResponse.ResultType}." +
                    $"\nErrorMessage: {sendMessageResponse.ErrorMessage}.");
            }

            await this.notificationService.UpdateSentNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: sendMessageResponse.TotalNumberOfSendThrottles,
                    statusCode: sendMessageResponse.StatusCode,
                    allSendStatusCodes: sendMessageResponse.AllSendStatusCodes,
                    errorMessage: sendMessageResponse.ErrorMessage,
                    reactionResult: "",
                    freeTextResult: "",
                    yesNoResult: "");

            // Throttled
            if (sendMessageResponse.ResultType == SendMessageResult.Throttled)
            {
                // Set send function throttled.
                await this.notificationService.SetSendNotificationThrottled(this.sendRetryDelayNumberOfSeconds);

                // Requeue.
                await this.sendQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                return;
            }
        }

        private async Task<IMessageActivity> GetMessageActivity(SendQueueMessageContent message)
        {
            var notification = await this.notificationRepo.GetAsync(
                NotificationDataTableNames.SendingNotificationsPartition,
                message.NotificationId);

            var parsedResult = AdaptiveCard.FromJson(notification.Content);
            var card = parsedResult.Card;

            //var recipentAadid = message.RecipientData.RecipientId;
            //var trackImageUrl = $"{this.appServiceUri}/api/GetUpdatedCard/Result?notificationId={message.NotificationId}&aadid={recipentAadid}.gif";

            //var pixel = new AdaptiveImage()
            //{
            //    Url = new Uri(trackImageUrl, UriKind.RelativeOrAbsolute),
            //    Spacing = AdaptiveSpacing.None,
            //    AltText = string.Empty,
            //};
            //pixel.AdditionalProperties.Add("width", "1px");
            //pixel.AdditionalProperties.Add("height", "1px");
            //card.Body.Add(pixel);
            this.teamsCard = card;
            this.aCard = card;
            notification.Content = notification.Content
                .Replace("\\n", "\\n\\r");
                //.Replace($"\"data\":{{\"notificationId\":{message.NotificationId}}}", "\"data\":{\"msteams\":{\"type\":\"task/fetch\"},\"data\":\"Invoke\"}");
            //var cardJson = new AdaptiveTextBlock()
            //{
            //    Text = notification.Content,
            //    Wrap = true,
            //};
            //card.Body.Add(cardJson);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCardContentType,
                //Content = JsonConvert.DeserializeObject(notification.Content),
                Content = card,
            };

            return MessageFactory.Attachment(adaptiveCardAttachment);
        }
    }
}
