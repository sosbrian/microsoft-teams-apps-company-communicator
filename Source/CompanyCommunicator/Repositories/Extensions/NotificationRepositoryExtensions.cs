// <copyright file="NotificationRepositoryExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Extensions for the repository of the notification data.
    /// </summary>
    public static class NotificationRepositoryExtensions
    {
        /// <summary>
        /// Create a new draft notification.
        /// </summary>
        /// <param name="notificationRepository">The notification repository.</param>
        /// <param name="notification">Draft Notification model class instance passed in from Web API.</param>
        /// <param name="userName">Name of the user who is running the application.</param>
        /// <returns>The newly created notification's id.</returns>
        public static async Task<string> CreateDraftNotificationAsync(
            this INotificationDataRepository notificationRepository,
            DraftNotification notification,
            string userName)
        {
            var newId = notificationRepository.TableRowKeyGenerator.CreateNewKeyOrderingOldestToMostRecent();

            var notificationEntity = new NotificationDataEntity
            {
                PartitionKey = NotificationDataTableNames.DraftNotificationsPartition,
                RowKey = newId,
                Id = newId,
                template = notification.Template,
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
                CreatedBy = userName,
                CreatedDate = DateTime.UtcNow,
                IsDraft = true,
                Teams = notification.Teams,
                Rosters = notification.Rosters,
                Groups = notification.Groups,
                UploadedList = notification.UploadedList,
                ExclusionList = notification.ExclusionList,
                AllUsers = notification.AllUsers,
            };

            await notificationRepository.CreateOrUpdateAsync(notificationEntity);

            return newId;
        }
    }
}
