// <copyright file="SurveyController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// Controller for saving survey result.
    /// </summary>
    [Route("api/Survey")]
    public class SurveyController : ControllerBase
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="SurveyController"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">Sent notification data repository instance.</param>
        public SurveyController(ISentNotificationDataRepository sentNotificationDataRepository)
        {
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

            // Update notification.
            notification.ReactionResult = reaction;
            notification.FreeTextResult = freetext;
            notification.YesNoResult = yesno;

            await this.sentNotificationDataRepository.InsertOrMergeAsync(notification);

            return this.Ok();
        }
    }
}