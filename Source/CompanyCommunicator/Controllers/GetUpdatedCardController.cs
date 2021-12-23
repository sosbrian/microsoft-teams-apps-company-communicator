// <copyright file="GetUpdatedCardController.cs company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

using Microsoft.Bot.Builder.Teams;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.CompanyCommunicator.Bot;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;

    /// <summary>
    /// Controller for saving survey result.
    /// </summary>
    [Route("api/GetUpdatedCard")]
    public class GetUpdatedCardController : ControllerBase
    {
        private readonly ISendingNotificationDataRepository notificationRepo;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly AdaptiveCardCreator adaptiveCardCreator;
        private readonly BotFrameworkHttpAdapter adapter;
        private readonly IBot authorBot;
        private readonly IBot userBot;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUpdatedCardController"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">Sent notification data repository instance.</param>
        /// <param name="notificationDataRepository">Sent notification data repository instance. Whatvever la. Who cares param.</param>
        /// <param name="notificationRepo">Sent notification data repository instance. Whatvever la. Who cares param. WFC.</param>
        public GetUpdatedCardController(
            INotificationDataRepository notificationDataRepository,
            ISendingNotificationDataRepository notificationRepo,
            AdaptiveCardCreator adaptiveCardCreator,
            ISentNotificationDataRepository sentNotificationDataRepository,
            CompanyCommunicatorBotAdapter adapter,
            AuthorTeamsActivityHandler authorBot,
            UserTeamsActivityHandler userBot)
        {
            this.notificationRepo = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.adaptiveCardCreator = adaptiveCardCreator ?? throw new ArgumentException(nameof(adaptiveCardCreator));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentException(nameof(notificationDataRepository)); //Get Card
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentException(nameof(sentNotificationDataRepository));
            this.adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));
            this.authorBot = authorBot ?? throw new ArgumentNullException(nameof(authorBot));
            this.userBot = userBot ?? throw new ArgumentNullException(nameof(userBot));
        }

        /// <summary>
        /// Get Survey Response.
        /// </summary>
        /// <param name="notificationId">Notification ID.</param>
        /// <param name="aadid">AAD ID of the user.</param>
        /// <returns>The result of an action method.</returns>
        [HttpGet("Result")]
        public async Task<IActionResult> PostSurveyResponse(
            [FromQuery] string notificationId,
            [FromQuery] string aadid)
        {
            var notification = await this.sentNotificationDataRepository.GetAsync(
                partitionKey: notificationId,
                rowKey: aadid);
            //var tempNotification = await this.notificationRepo.GetAsync(
            //    NotificationDataTableNames.SendingNotificationsPartition,
            //    notificationId);
            //var entityNotification = await this.sentNotificationDataRepository.GetAsync(
            //    NotificationDataTableNames.SentNotificationsPartition,
            //    notificationId);
            var textNotification = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                notificationId);
            var vCard = this.adaptiveCardCreator.CreateAdaptiveCard(textNotification, true);
            var test = vCard.ToJson()
                .Replace("\"version\": \"1.2\",", "\"version\": \"1.0\",")
                .Replace("\\n", "\\n\\r");
                //.Replace("\r\n", string.Empty);

            if ((textNotification.SurReaction == true && notification.ReactionResult == string.Empty)
                || (textNotification.SurFreeText == true && notification.FreeTextResult == string.Empty)
                || (textNotification.SurYesNo == true && notification.YesNoResult == string.Empty))
            {
                return this.NoContent();
            }

            //this.Response.Headers.Add("CARD-UPDATE-IN-BODY", "true");

            // await this.sentNotificationDataRepository.InsertOrMergeAsync(notification);

            this.Response.Headers.Add("CARD-UPDATE-IN-BODY", "true");
            return this.Ok(test);
        }

        [HttpGet("TeamsResult")]
        public async Task<IActionResult> PostTeamsSurveyResponse(
            [FromQuery] string notificationId,
            [FromQuery] string aadid)
        {
            return this.NoContent();
        }
    }
}