﻿// <copyright file="DraftNotificationsController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.DraftNotificationPreview;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;

    /// <summary>
    /// Controller for the draft notification data.
    /// </summary>
    [Route("api/draftNotifications")]
    //[Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class DraftNotificationsController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ITeamDataRepository teamDataRepository;
        private readonly IDraftNotificationPreviewService draftNotificationPreviewService;
        private readonly IGroupsService groupsService;
        private readonly IAppSettingsService appSettingsService;
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="DraftNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository instance.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        /// <param name="draftNotificationPreviewService">Draft notification preview service.</param>
        /// <param name="appSettingsService">App Settings service.</param>
        /// <param name="localizer">Localization service.</param>
        /// <param name="groupsService">group service.</param>
        public DraftNotificationsController(
            INotificationDataRepository notificationDataRepository,
            ITeamDataRepository teamDataRepository,
            IDraftNotificationPreviewService draftNotificationPreviewService,
            IAppSettingsService appSettingsService,
            IStringLocalizer<Strings> localizer,
            IGroupsService groupsService)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
            this.draftNotificationPreviewService = draftNotificationPreviewService ?? throw new ArgumentNullException(nameof(draftNotificationPreviewService));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.groupsService = groupsService ?? throw new ArgumentNullException(nameof(groupsService));
            this.appSettingsService = appSettingsService ?? throw new ArgumentNullException(nameof(appSettingsService));
        }

        /// <summary>
        /// Create a new draft notification.
        /// </summary>
        /// <param name="notification">A new Draft Notification to be created.</param>
        /// <returns>The created notification's id.</returns>
        [HttpPost]
        public async Task<ActionResult<string>> CreateDraftNotificationAsync([FromBody] DraftNotification notification)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            if (!notification.Validate(this.localizer, out string errorMessage))
            {
                return this.BadRequest(errorMessage);
            }

            var containsHiddenMembership = await this.groupsService.ContainsHiddenMembershipAsync(notification.Groups);
            if (containsHiddenMembership)
            {
                return this.Forbid();
            }

            var notificationId = await this.notificationDataRepository.CreateDraftNotificationAsync(
                notification,
                this.HttpContext.User?.Identity?.Name);
            return this.Ok(notificationId);
        }

        /// <summary>
        /// Duplicate an existing draft notification.
        /// </summary>
        /// <param name="id">The id of a Draft Notification to be duplicated.</param>
        /// <returns>If the passed in id is invalid, it returns 404 not found error. Otherwise, it returns 200 OK.</returns>
        [HttpPost("duplicates/{id}")]
        public async Task<IActionResult> DuplicateDraftNotificationAsync(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            var notificationEntity = await this.FindNotificationToDuplicate(id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var createdBy = this.HttpContext.User?.Identity?.Name;
            notificationEntity.Title = this.localizer.GetString("DuplicateText", notificationEntity.Title);
            await this.notificationDataRepository.DuplicateDraftNotificationAsync(notificationEntity, createdBy);

            return this.Ok();
        }

        /// <summary>
        /// Update an existing draft notification.
        /// </summary>
        /// <param name="notification">An existing Draft Notification to be updated.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [HttpPut]
        public async Task<IActionResult> UpdateDraftNotificationAsync([FromBody] DraftNotification notification)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            var containsHiddenMembership = await this.groupsService.ContainsHiddenMembershipAsync(notification.Groups);
            if (containsHiddenMembership)
            {
                return this.Forbid();
            }

            if (!notification.Validate(this.localizer, out string errorMessage))
            {
                return this.BadRequest(errorMessage);
            }

            var notificationEntity = new NotificationDataEntity
            {
                PartitionKey = NotificationDataTableNames.DraftNotificationsPartition,
                RowKey = notification.Id,
                Id = notification.Id,
                template = notification.Template,
                PriLanguage = notification.PriLanguage,
                SecLanguage = notification.SecLanguage,
                SenderTemplate = notification.SenderTemplate,
                Title = notification.Title,
                ImageLink = notification.ImageLink,
                VideoLink = notification.VideoLink,
                Summary = notification.Summary,
                Alignment = notification.Alignment,
                BoldSummary = notification.BoldSummary,
                FontSummary = notification.FontSummary,
                FontSizeSummary = notification.FontSizeSummary,
                FontColorSummary = notification.FontColorSummary,
                Author = notification.Author,
                ButtonTitle = notification.ButtonTitle,
                ButtonLink = notification.ButtonLink,
                ButtonTitle2 = notification.ButtonTitle2,
                ButtonLink2 = notification.ButtonLink2,
                ButtonTitle3 = notification.ButtonTitle3,
                ButtonLink3 = notification.ButtonLink3,
                ButtonTitle4 = notification.ButtonTitle4,
                ButtonLink4 = notification.ButtonLink4,
                ButtonTitle5 = notification.ButtonTitle5,
                ButtonLink5 = notification.ButtonLink5,
                SurReaction = notification.SurReaction,
                ReactionQuestion = notification.ReactionQuestion,
                SurFreeText = notification.SurFreeText,
                FreeTextQuestion = notification.FreeTextQuestion,
                SurYesNo = notification.SurYesNo,
                YesNoQuestion = notification.YesNoQuestion,
                SurLinkToSurvey = notification.SurLinkToSurvey,
                LinkToSurvey = notification.LinkToSurvey,
                SecSenderTemplate = notification.SecSenderTemplate,
                SecTitle = notification.SecTitle,
                SecImageLink = notification.SecImageLink,
                SecVideoLink = notification.SecVideoLink,
                SecSummary = notification.SecSummary,
                SecAlignment = notification.SecAlignment,
                SecBoldSummary = notification.SecBoldSummary,
                SecFontSummary = notification.SecFontSummary,
                SecFontSizeSummary = notification.SecFontSizeSummary,
                SecFontColorSummary = notification.SecFontColorSummary,
                SecAuthor = notification.SecAuthor,
                SecButtonTitle = notification.SecButtonTitle,
                SecButtonLink = notification.SecButtonLink,
                SecButtonTitle2 = notification.SecButtonTitle2,
                SecButtonLink2 = notification.SecButtonLink2,
                SecButtonTitle3 = notification.SecButtonTitle3,
                SecButtonLink3 = notification.SecButtonLink3,
                SecButtonTitle4 = notification.SecButtonTitle4,
                SecButtonLink4 = notification.SecButtonLink4,
                SecButtonTitle5 = notification.SecButtonTitle5,
                SecButtonLink5 = notification.SecButtonLink5,
                SecSurReaction = notification.SecSurReaction,
                SecReactionQuestion = notification.SecReactionQuestion,
                SecSurFreeText = notification.SecSurFreeText,
                SecFreeTextQuestion = notification.SecFreeTextQuestion,
                SecSurYesNo = notification.SecSurYesNo,
                SecYesNoQuestion = notification.SecYesNoQuestion,
                SecSurLinkToSurvey = notification.SecSurLinkToSurvey,
                SecLinkToSurvey = notification.SecLinkToSurvey,
                CreatedBy = this.HttpContext.User?.Identity?.Name,
                CreatedDate = DateTime.UtcNow,
                IsDraft = true,
                Teams = notification.Teams,
                Rosters = notification.Rosters,
                Groups = notification.Groups,
                UploadedList = notification.UploadedList,
                UploadedListName = notification.UploadedListName,
                EmailOption = notification.EmailOption,
                ExclusionList = notification.ExclusionList,
                AllUsers = notification.AllUsers,
                IsScheduled = notification.IsScheduled,
                ScheduledDate = notification.ScheduledDate,
                IsExpirySet = notification.IsExpirySet,
                ExpiryDate = notification.ExpiryDate,
                IsExpiredContentErased = notification.IsExpiredContentErased,
            };

            await this.notificationDataRepository.CreateOrUpdateAsync(notificationEntity);
            return this.Ok();
        }

        /// <summary>
        /// Delete an existing draft notification.
        /// </summary>
        /// <param name="id">The id of the draft notification to be deleted.</param>
        /// <returns>If the passed in Id is invalid, it returns 404 not found error. Otherwise, it returns 200 OK.</returns>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteDraftNotificationAsync(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            await this.notificationDataRepository.DeleteAsync(notificationEntity);
            return this.Ok();
        }

        /// <summary>
        /// Get draft notifications.
        /// </summary>
        /// <returns>A list of <see cref="DraftNotificationSummary"/> instances.</returns>
        [HttpGet]
        public async Task<ActionResult<IEnumerable<DraftNotificationSummary>>> GetAllDraftNotificationsAsync()
        {
            var notificationEntities = await this.notificationDataRepository.GetAllDraftNotificationsAsync();

            var result = new List<DraftNotificationSummary>();
            foreach (var notificationEntity in notificationEntities)
            {
                var summary = new DraftNotificationSummary
                {
                    Id = notificationEntity.Id,
                    Title = notificationEntity.Title,
                };

                result.Add(summary);
            }

            return result;
        }

        /// <summary>
        /// Get scheduled notifications. Those are draft notifications with a scheduledate.
        /// </summary>
        /// <returns>A list of <see cref="DraftNotificationSummary"/> instances.</returns>
        [HttpGet("scheduled")]
        public async Task<ActionResult<IEnumerable<DraftNotificationSummary>>> GetAllScheduledNotificationsAsync()
        {
            var notificationEntities = await this.notificationDataRepository.GetAllScheduledNotificationsAsync();

            var result = new List<DraftNotificationSummary>();
            foreach (var notificationEntity in notificationEntities)
            {
                var summary = new DraftNotificationSummary
                {
                    Id = notificationEntity.Id,
                    Title = notificationEntity.Title,
                    ScheduledDate = notificationEntity.ScheduledDate,
                    ExpiryDate = notificationEntity.ExpiryDate,
                };

                result.Add(summary);
            }

            // sorts the scheduled messages by date from the most recent
            result.Sort((r1, r2) => r1.ScheduledDate.Value.CompareTo(r2.ScheduledDate.Value));
            return result;
        }

        /// <summary>
        /// Get a draft notification by Id.
        /// </summary>
        /// <param name="id">Draft notification Id.</param>
        /// <returns>It returns the draft notification with the passed in id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("{id}")]
        public async Task<ActionResult<DraftNotification>> GetDraftNotificationByIdAsync(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var result = new DraftNotification
            {
                Id = notificationEntity.Id,
                Template = notificationEntity.template,
                PriLanguage = notificationEntity.PriLanguage,
                SecLanguage = notificationEntity.SecLanguage,
                SenderTemplate = notificationEntity.SenderTemplate,
                Title = notificationEntity.Title,
                ImageLink = notificationEntity.ImageLink,
                VideoLink = notificationEntity.VideoLink,
                Summary = notificationEntity.Summary,
                Alignment = notificationEntity.Alignment,
                BoldSummary = notificationEntity.BoldSummary,
                FontSummary = notificationEntity.FontSummary,
                FontSizeSummary = notificationEntity.FontSizeSummary,
                FontColorSummary = notificationEntity.FontColorSummary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                ButtonTitle2 = notificationEntity.ButtonTitle2,
                ButtonLink2 = notificationEntity.ButtonLink2,
                ButtonTitle3 = notificationEntity.ButtonTitle3,
                ButtonLink3 = notificationEntity.ButtonLink3,
                ButtonTitle4 = notificationEntity.ButtonTitle4,
                ButtonLink4 = notificationEntity.ButtonLink4,
                ButtonTitle5 = notificationEntity.ButtonTitle5,
                ButtonLink5 = notificationEntity.ButtonLink5,
                SurReaction = notificationEntity.SurReaction,
                ReactionQuestion = notificationEntity.ReactionQuestion,
                SurFreeText = notificationEntity.SurFreeText,
                FreeTextQuestion = notificationEntity.FreeTextQuestion,
                SurYesNo = notificationEntity.SurYesNo,
                YesNoQuestion = notificationEntity.YesNoQuestion,
                SurLinkToSurvey = notificationEntity.SurLinkToSurvey,
                LinkToSurvey = notificationEntity.LinkToSurvey,
                SecSenderTemplate = notificationEntity.SecSenderTemplate,
                SecTitle = notificationEntity.SecTitle,
                SecImageLink = notificationEntity.SecImageLink,
                SecVideoLink = notificationEntity.SecVideoLink,
                SecSummary = notificationEntity.SecSummary,
                SecAlignment = notificationEntity.SecAlignment,
                SecBoldSummary = notificationEntity.SecBoldSummary,
                SecFontSummary = notificationEntity.SecFontSummary,
                SecFontSizeSummary = notificationEntity.SecFontSizeSummary,
                SecFontColorSummary = notificationEntity.SecFontColorSummary,
                SecAuthor = notificationEntity.SecAuthor,
                SecButtonTitle = notificationEntity.SecButtonTitle,
                SecButtonLink = notificationEntity.SecButtonLink,
                SecButtonTitle2 = notificationEntity.SecButtonTitle2,
                SecButtonLink2 = notificationEntity.SecButtonLink2,
                SecButtonTitle3 = notificationEntity.SecButtonTitle3,
                SecButtonLink3 = notificationEntity.SecButtonLink3,
                SecButtonTitle4 = notificationEntity.SecButtonTitle4,
                SecButtonLink4 = notificationEntity.SecButtonLink4,
                SecButtonTitle5 = notificationEntity.SecButtonTitle5,
                SecButtonLink5 = notificationEntity.SecButtonLink5,
                SecSurReaction = notificationEntity.SecSurReaction,
                SecReactionQuestion = notificationEntity.SecReactionQuestion,
                SecSurFreeText = notificationEntity.SecSurFreeText,
                SecFreeTextQuestion = notificationEntity.SecFreeTextQuestion,
                SecSurYesNo = notificationEntity.SecSurYesNo,
                SecYesNoQuestion = notificationEntity.SecYesNoQuestion,
                SecSurLinkToSurvey = notificationEntity.SecSurLinkToSurvey,
                SecLinkToSurvey = notificationEntity.SecLinkToSurvey,
                CreatedDateTime = notificationEntity.CreatedDate,
                Teams = notificationEntity.Teams,
                Rosters = notificationEntity.Rosters,
                Groups = notificationEntity.Groups,
                UploadedList = notificationEntity.UploadedList,
                UploadedListName = notificationEntity.UploadedListName,
                EmailOption = notificationEntity.EmailOption,
                ExclusionList = notificationEntity.ExclusionList,
                AllUsers = notificationEntity.AllUsers,
                IsScheduled = notificationEntity.IsScheduled,
                ScheduledDate = notificationEntity.ScheduledDate,
                IsExpirySet = notificationEntity.IsExpirySet,
                ExpiryDate = notificationEntity.ExpiryDate,
                IsExpiredContentErased = notificationEntity.IsExpiredContentErased,
            };

            return this.Ok(result);
        }

        /// <summary>
        /// Get draft notification summary (for consent page) by notification Id.
        /// </summary>
        /// <param name="notificationId">Draft notification Id.</param>
        /// <returns>It returns the draft notification summary (for consent page) with the passed in id.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("consentSummaries/{notificationId}")]
        public async Task<ActionResult<DraftNotificationSummaryForConsent>> GetDraftNotificationSummaryForConsentByIdAsync(string notificationId)
        {
            if (notificationId == null)
            {
                throw new ArgumentNullException(nameof(notificationId));
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                notificationId);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var groupNames = await this.groupsService
                .GetByIdsAsync(notificationEntity.Groups)
                .Select(x => x.DisplayName).
                ToListAsync();

            var result = new DraftNotificationSummaryForConsent
            {
                NotificationId = notificationId,
                TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Teams),
                RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Rosters),
                GroupNames = groupNames,
                AllUsers = notificationEntity.AllUsers,
                UploadedList = notificationEntity.UploadedList,
                UploadedListName = notificationEntity.UploadedListName,
                EmailOption = notificationEntity.EmailOption,
                ExclusionList = notificationEntity.ExclusionList,
            };

            return this.Ok(result);
        }

        /// <summary>
        /// Preview draft notification.
        /// </summary>
        /// <param name="draftNotificationPreviewRequest">Draft notification preview request.</param>
        /// <returns>
        /// It returns 400 bad request error if the incoming parameter, draftNotificationPreviewRequest, is invalid.
        /// It returns 404 not found error if the DraftNotificationId or TeamsTeamId (contained in draftNotificationPreviewRequest) is not found in the table storage.
        /// It returns 500 internal error if this method throws an unhandled exception.
        /// It returns 429 too many requests error if the preview request is throttled by the bot service.
        /// It returns 200 OK if the method is executed successfully.</returns>
        [HttpPost("previews")]
        public async Task<ActionResult> PreviewDraftNotificationAsync(
            [FromBody] DraftNotificationPreviewRequest draftNotificationPreviewRequest)
        {
            if (draftNotificationPreviewRequest == null
                || string.IsNullOrWhiteSpace(draftNotificationPreviewRequest.DraftNotificationId)
                || string.IsNullOrWhiteSpace(draftNotificationPreviewRequest.TeamsTeamId)
                || string.IsNullOrWhiteSpace(draftNotificationPreviewRequest.TeamsChannelId))
            {
                return this.BadRequest();
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                draftNotificationPreviewRequest.DraftNotificationId);
            if (notificationEntity == null)
            {
                return this.BadRequest($"Notification {draftNotificationPreviewRequest.DraftNotificationId} not found.");
            }

            var teamDataEntity = new TeamDataEntity();
            teamDataEntity.TenantId = this.HttpContext.User.FindFirstValue(Common.Constants.ClaimTypeTenantId);
            teamDataEntity.ServiceUrl = await this.appSettingsService.GetServiceUrlAsync();
            var result = await this.draftNotificationPreviewService.SendPreview(
                notificationEntity,
                teamDataEntity,
                draftNotificationPreviewRequest.TeamsChannelId);
            return this.StatusCode((int)result);
        }

        private async Task<NotificationDataEntity> FindNotificationToDuplicate(string notificationId)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                notificationId);
            if (notificationEntity == null)
            {
                notificationEntity = await this.notificationDataRepository.GetAsync(
                    NotificationDataTableNames.SentNotificationsPartition,
                    notificationId);
            }

            return notificationEntity;
        }
    }
}
