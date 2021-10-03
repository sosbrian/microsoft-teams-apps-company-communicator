// <copyright file="NotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the notification data in the table storage.
    /// </summary>
    public class NotificationDataRepository : BaseRepository<NotificationDataEntity>, INotificationDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        /// <param name="tableRowKeyGenerator">Table row key generator service.</param>
        public NotificationDataRepository(
            ILogger<NotificationDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions,
            TableRowKeyGenerator tableRowKeyGenerator)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: NotificationDataTableNames.TableName,
                  defaultPartitionKey: NotificationDataTableNames.DraftNotificationsPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
            this.TableRowKeyGenerator = tableRowKeyGenerator;
        }

        /// <inheritdoc/>
        public TableRowKeyGenerator TableRowKeyGenerator { get; }

        /// <inheritdoc/>
        public async Task<IEnumerable<NotificationDataEntity>> GetAllDraftNotificationsAsync()
        {
            var result = await this.GetAllAsync(NotificationDataTableNames.DraftNotificationsPartition);

            return result;
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<NotificationDataEntity>> GetMostRecentSentNotificationsAsync()
        {
            var result = await this.GetAllAsync(NotificationDataTableNames.SentNotificationsPartition, 25);

            return result;
        }

        /// <inheritdoc/>
        public async Task<string> MoveDraftToSentPartitionAsync(NotificationDataEntity draftNotificationEntity)
        {
            try
            {
                if (draftNotificationEntity == null)
                {
                    throw new ArgumentNullException(nameof(draftNotificationEntity));
                }

                var newSentNotificationId = this.TableRowKeyGenerator.CreateNewKeyOrderingMostRecentToOldest();

                // Create a sent notification based on the draft notification.
                var sentNotificationEntity = new NotificationDataEntity
                {
                    PartitionKey = NotificationDataTableNames.SentNotificationsPartition,
                    RowKey = newSentNotificationId,
                    Id = newSentNotificationId,
                    template = draftNotificationEntity.template,
                    SenderTemplate = draftNotificationEntity.SenderTemplate,
                    Title = draftNotificationEntity.Title,
                    ImageLink = draftNotificationEntity.ImageLink,
                    VideoLink = draftNotificationEntity.VideoLink,
                    Summary = draftNotificationEntity.Summary,
                    Alignment = draftNotificationEntity.Alignment,
                    BoldSummary = draftNotificationEntity.BoldSummary,
                    FontSummary = draftNotificationEntity.FontSummary,
                    FontSizeSummary = draftNotificationEntity.FontSizeSummary,
                    FontColorSummary = draftNotificationEntity.FontColorSummary,
                    Author = draftNotificationEntity.Author,
                    ButtonTitle = draftNotificationEntity.ButtonTitle,
                    ButtonLink = draftNotificationEntity.ButtonLink,
                    ButtonTitle2 = draftNotificationEntity.ButtonTitle2,
                    ButtonLink2 = draftNotificationEntity.ButtonLink2,
                    ButtonTitle3 = draftNotificationEntity.ButtonTitle3,
                    ButtonLink3 = draftNotificationEntity.ButtonLink3,
                    ButtonTitle4 = draftNotificationEntity.ButtonTitle4,
                    ButtonLink4 = draftNotificationEntity.ButtonLink4,
                    ButtonTitle5 = draftNotificationEntity.ButtonTitle5,
                    ButtonLink5 = draftNotificationEntity.ButtonLink5,
                    SurReaction = draftNotificationEntity.SurReaction,
                    ReactionQuestion = draftNotificationEntity.ReactionQuestion,
                    SurFreeText = draftNotificationEntity.SurFreeText,
                    FreeTextQuestion = draftNotificationEntity.FreeTextQuestion,
                    SurYesNo = draftNotificationEntity.SurYesNo,
                    YesNoQuestion = draftNotificationEntity.YesNoQuestion,
                    SurLinkToSurvey = draftNotificationEntity.SurLinkToSurvey,
                    LinkToSurvey = draftNotificationEntity.LinkToSurvey,
                    CreatedBy = draftNotificationEntity.CreatedBy,
                    CreatedDate = draftNotificationEntity.CreatedDate,
                    SentDate = null,
                    IsDraft = false,
                    Teams = draftNotificationEntity.Teams,
                    Rosters = draftNotificationEntity.Rosters,
                    Groups = draftNotificationEntity.Groups,
                    AllUsers = draftNotificationEntity.AllUsers,
                    UploadedList = draftNotificationEntity.UploadedList,
                    UploadedListName = draftNotificationEntity.UploadedListName,
                    ExclusionList = draftNotificationEntity.ExclusionList,
                    MessageVersion = draftNotificationEntity.MessageVersion,
                    Succeeded = 0,
                    Failed = 0,
                    Throttled = 0,
                    TotalMessageCount = draftNotificationEntity.TotalMessageCount,
                    SendingStartedDate = DateTime.UtcNow,
                    Status = NotificationStatus.Queued.ToString(),
                };
                await this.CreateOrUpdateAsync(sentNotificationEntity);

                // Delete the draft notification.
                await this.DeleteAsync(draftNotificationEntity);

                return newSentNotificationId;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task DuplicateDraftNotificationAsync(
            NotificationDataEntity notificationEntity,
            string createdBy)
        {
            try
            {
                var newId = this.TableRowKeyGenerator.CreateNewKeyOrderingOldestToMostRecent();

                // TODO: Set the string "(copy)" in a resource file for multi-language support.
                var newNotificationEntity = new NotificationDataEntity
                {
                    PartitionKey = NotificationDataTableNames.DraftNotificationsPartition,
                    RowKey = newId,
                    Id = newId,
                    template = notificationEntity.template,
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
                    CreatedBy = createdBy,
                    CreatedDate = DateTime.UtcNow,
                    IsDraft = true,
                    Teams = notificationEntity.Teams,
                    Groups = notificationEntity.Groups,
                    Rosters = notificationEntity.Rosters,
                    UploadedList = notificationEntity.UploadedList,
                    UploadedListName = notificationEntity.UploadedListName,
                    ExclusionList = notificationEntity.ExclusionList,
                    AllUsers = notificationEntity.AllUsers,
                };

                await this.CreateOrUpdateAsync(newNotificationEntity);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task UpdateNotificationStatusAsync(string notificationId, NotificationStatus status)
        {
            var notificationDataEntity = await this.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                notificationId);

            if (notificationDataEntity != null)
            {
                notificationDataEntity.Status = status.ToString();
                await this.CreateOrUpdateAsync(notificationDataEntity);
            }
        }

        /// <inheritdoc/>
        public async Task SaveExceptionInNotificationDataEntityAsync(
            string notificationDataEntityId,
            string errorMessage)
        {
            var notificationDataEntity = await this.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                notificationDataEntityId);
            if (notificationDataEntity != null)
            {
                notificationDataEntity.ErrorMessage =
                    this.AppendNewLine(notificationDataEntity.ErrorMessage, errorMessage);
                notificationDataEntity.Status = NotificationStatus.Failed.ToString();

                // Set the end date as current date.
                notificationDataEntity.SentDate = DateTime.UtcNow;

                await this.CreateOrUpdateAsync(notificationDataEntity);
            }
        }

        /// <inheritdoc/>
        public async Task SaveWarningInNotificationDataEntityAsync(
            string notificationDataEntityId,
            string warningMessage)
        {
            try
            {
                var notificationDataEntity = await this.GetAsync(
                    NotificationDataTableNames.SentNotificationsPartition,
                    notificationDataEntityId);
                if (notificationDataEntity != null)
                {
                    notificationDataEntity.WarningMessage =
                        this.AppendNewLine(notificationDataEntity.WarningMessage, warningMessage);
                    await this.CreateOrUpdateAsync(notificationDataEntity);
                }
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        private string AppendNewLine(string originalString, string newString)
        {
            return string.IsNullOrWhiteSpace(originalString)
                ? newString
                : $"{originalString}{Environment.NewLine}{newString}";
        }
    }
}
