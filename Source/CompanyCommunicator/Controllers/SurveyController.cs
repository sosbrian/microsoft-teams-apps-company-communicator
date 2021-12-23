// <copyright file="SurveyController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

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
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Microsoft.Bot.Builder;

    /// <summary>
    /// Controller for saving survey result.
    /// </summary>
    [Route("api/Survey")]
    public class SurveyController : ControllerBase
    {
        private readonly ISendingNotificationDataRepository notificationRepo;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly AdaptiveCardCreator adaptiveCardCreator;

        /// <summary>
        /// Initializes a new instance of the <see cref="SurveyController"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">Sent notification data repository instance.</param>
        /// <param name="notificationDataRepository">Sent notification data repository instance. Whatvever la. Who cares param.</param>
        /// <param name="notificationRepo">Sent notification data repository instance. Whatvever la. Who cares param. WFC.</param>
        /// <param name="adaptiveCardCreator">Create Adaptive card.</param>
        public SurveyController(
            INotificationDataRepository notificationDataRepository,
            ISendingNotificationDataRepository notificationRepo,
            AdaptiveCardCreator adaptiveCardCreator,
            ISentNotificationDataRepository sentNotificationDataRepository)
        {
            this.notificationRepo = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.adaptiveCardCreator = adaptiveCardCreator ?? throw new ArgumentException(nameof(adaptiveCardCreator));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentException(nameof(notificationDataRepository)); //Get Card
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentException(nameof(sentNotificationDataRepository));
        }

        /// <summary>
        /// Get Survey Response.
        /// </summary>
        /// <param name="notificationId">Notification ID.</param>
        /// <param name="aadid">AAD ID of the user.</param>
        /// <param name="reaction">Reaction Survey Response.</param>
        /// <param name="freetext">Free Text Survey Response.</param>
        /// <param name="yesno">Yes/No Survey Response.</param>
        /// <returns>The result of an action method.</returns>
        [HttpGet("Result")]
        public async Task<IActionResult> PostSurveyResponse(
            [FromQuery] string notificationId,
            [FromQuery] string aadid,
            [FromQuery] string reaction,
            [FromQuery] string freetext,
            [FromQuery] string yesno)
        {
            var notification = await this.sentNotificationDataRepository.GetAsync(
                partitionKey: notificationId,
                rowKey: aadid);
            //var tempNotification = await this.notificationRepo.GetAsync(
            //    NotificationDataTableNames.SendingNotificationsPartition,
            //    notificationId);
            var textNotification = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                notificationId);
            var vCard = this.adaptiveCardCreator.CreateAdaptiveCard(textNotification, true);
            var test = vCard.ToJson()
                .Replace("\"version\": \"1.2\",", "\"version\": \"1.0\",")
                .Replace("\\n", "\\n\\r")
                .Replace("\r\n", string.Empty);
            // if(textNotification.)
            if ((textNotification.SurReaction == true && reaction == "{{Reaction.value}}")
                || (textNotification.SurFreeText == true && freetext == null)
                || (textNotification.SurYesNo == true && yesno == "{{YesNo.value}}"))
            {
                return this.BadRequest("Result Not Found.");
            }

            //if (textNotification.SurFreeText == true && freetext == null)
            //{
            //    return this.NotFound("Result Not Found.");
            //}

            //if (textNotification.SurYesNo == true && yesno == "{{YesNo.value}}")
            //{
            //    return this.NotFound("Result Not Found.");
            //}

            //if (reaction == "{{}}" || freetext == "{{}}" || yesno == "{{}}")
            //{
            //    return this.NotFound("Result Not Found.");
            //}

            // Update notification.
            notification.ReactionResult = reaction;
            notification.FreeTextResult = freetext;
            notification.YesNoResult = yesno;

            this.Response.Headers.Add("CARD-UPDATE-IN-BODY", "true");

            await this.sentNotificationDataRepository.InsertOrMergeAsync(notification);

            return this.Ok(test);
        }
    }
}