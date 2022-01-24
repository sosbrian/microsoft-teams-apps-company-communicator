// <copyright file="NotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// Repository of the notification data in the table storage.
    /// </summary>
    public class NotificationDataRepository : BaseRepository<NotificationDataEntity>, INotificationDataRepository
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;

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
            string strFilter = TableQuery.GenerateFilterConditionForBool("IsScheduled", QueryComparisons.Equal, false);
            var result = await this.GetWithFilterAsync(strFilter, NotificationDataTableNames.DraftNotificationsPartition);

            return result;
        }

        public async Task<IEnumerable<NotificationDataEntity>> GetAllScheduledNotificationsAsync()
        {
            string strFilter = TableQuery.GenerateFilterConditionForBool("IsScheduled", QueryComparisons.Equal, true);
            var result = await this.GetWithFilterAsync(strFilter, NotificationDataTableNames.DraftNotificationsPartition);

            return result;
        }

        public async Task<IEnumerable<NotificationDataEntity>> GetAllPendingScheduledNotificationsAsync()
        {
            DateTime now = DateTime.UtcNow;

            string filter1 = TableQuery.GenerateFilterConditionForBool("IsScheduled", QueryComparisons.Equal, true);
            string filter2 = TableQuery.GenerateFilterConditionForDate("ScheduledDate", QueryComparisons.LessThanOrEqual, now);
            string filter = TableQuery.CombineFilters(filter1, TableOperators.And, filter2);

            var result = await this.GetWithFilterAsync(filter, NotificationDataTableNames.DraftNotificationsPartition);

            return result;
        }

        public async Task<IEnumerable<NotificationDataEntity>> GetNonErasedExpiredNotificationsAsync()
        {
            DateTime now = DateTime.UtcNow;
            DateTime pastOneDay = DateTime.UtcNow.AddHours(-24);


            string filter1 = TableQuery.GenerateFilterConditionForBool("IsExpiredContentErased", QueryComparisons.NotEqual, true);
            string filter2 = TableQuery.GenerateFilterConditionForDate("ExpiryDate", QueryComparisons.LessThanOrEqual, now);
            string filterA = TableQuery.CombineFilters(filter1, TableOperators.And, filter2);

            string filterB = TableQuery.GenerateFilterConditionForBool("IsExpirySet", QueryComparisons.Equal, true);

            string filterC = TableQuery.CombineFilters(filterA, TableOperators.And, filterB);

            string filter5 = TableQuery.GenerateFilterConditionForDate("ExpiryDate", QueryComparisons.LessThanOrEqual, now);
            string filter6 = TableQuery.GenerateFilterConditionForDate("ExpiryDate", QueryComparisons.GreaterThanOrEqual, pastOneDay);

            string filterD = TableQuery.CombineFilters(filter5, TableOperators.And, filter6);

            string filter9 = TableQuery.GenerateFilterConditionForBool("IsExpirySet", QueryComparisons.Equal, true);

            string filterE = TableQuery.CombineFilters(filterD, TableOperators.And, filter9);

            string filterFinal = TableQuery.CombineFilters(filterC, TableOperators.Or, filterE);

            var result = await this.GetWithFilterAsync(filterFinal, NotificationDataTableNames.SentNotificationsPartition);

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
                    PriLanguage = draftNotificationEntity.PriLanguage,
                    SecLanguage = draftNotificationEntity.SecLanguage,
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
                    SecSenderTemplate = draftNotificationEntity.SecSenderTemplate,
                    SecTitle = draftNotificationEntity.SecTitle,
                    SecImageLink = draftNotificationEntity.SecImageLink,
                    SecVideoLink = draftNotificationEntity.SecVideoLink,
                    SecSummary = draftNotificationEntity.SecSummary,
                    SecAlignment = draftNotificationEntity.SecAlignment,
                    SecBoldSummary = draftNotificationEntity.SecBoldSummary,
                    SecFontSummary = draftNotificationEntity.SecFontSummary,
                    SecFontSizeSummary = draftNotificationEntity.SecFontSizeSummary,
                    SecFontColorSummary = draftNotificationEntity.SecFontColorSummary,
                    SecAuthor = draftNotificationEntity.SecAuthor,
                    SecButtonTitle = draftNotificationEntity.SecButtonTitle,
                    SecButtonLink = draftNotificationEntity.SecButtonLink,
                    SecButtonTitle2 = draftNotificationEntity.SecButtonTitle2,
                    SecButtonLink2 = draftNotificationEntity.SecButtonLink2,
                    SecButtonTitle3 = draftNotificationEntity.SecButtonTitle3,
                    SecButtonLink3 = draftNotificationEntity.SecButtonLink3,
                    SecButtonTitle4 = draftNotificationEntity.SecButtonTitle4,
                    SecButtonLink4 = draftNotificationEntity.SecButtonLink4,
                    SecButtonTitle5 = draftNotificationEntity.SecButtonTitle5,
                    SecButtonLink5 = draftNotificationEntity.SecButtonLink5,
                    SecSurReaction = draftNotificationEntity.SecSurReaction,
                    SecReactionQuestion = draftNotificationEntity.SecReactionQuestion,
                    SecSurFreeText = draftNotificationEntity.SecSurFreeText,
                    SecFreeTextQuestion = draftNotificationEntity.SecFreeTextQuestion,
                    SecSurYesNo = draftNotificationEntity.SecSurYesNo,
                    SecYesNoQuestion = draftNotificationEntity.SecYesNoQuestion,
                    SecSurLinkToSurvey = draftNotificationEntity.SecSurLinkToSurvey,
                    SecLinkToSurvey = draftNotificationEntity.SecLinkToSurvey,
                    CreatedBy = draftNotificationEntity.CreatedBy,
                    CreatedDate = draftNotificationEntity.CreatedDate,
                    IsScheduled = draftNotificationEntity.IsScheduled,
                    ScheduledDate = draftNotificationEntity.ScheduledDate,
                    IsExpirySet = draftNotificationEntity.IsExpirySet,
                    ExpiryDate = draftNotificationEntity.ExpiryDate,
                    IsExpiredContentErased = draftNotificationEntity.IsExpiredContentErased,
                    SentDate = null,
                    IsDraft = false,
                    Teams = draftNotificationEntity.Teams,
                    Rosters = draftNotificationEntity.Rosters,
                    Groups = draftNotificationEntity.Groups,
                    AllUsers = draftNotificationEntity.AllUsers,
                    UploadedList = draftNotificationEntity.UploadedList,
                    UploadedListName = draftNotificationEntity.UploadedListName,
                    EmailOption = draftNotificationEntity.EmailOption,
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
                    PriLanguage = notificationEntity.PriLanguage,
                    SecLanguage = notificationEntity.SecLanguage,
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
                    CreatedBy = createdBy,
                    CreatedDate = DateTime.UtcNow,
                    IsScheduled = false,
                    ScheduledDate = null,
                    IsExpirySet = false,
                    ExpiryDate = null,
                    IsExpiredContentErased = false,
                    IsDraft = true,
                    Teams = notificationEntity.Teams,
                    Groups = notificationEntity.Groups,
                    Rosters = notificationEntity.Rosters,
                    UploadedList = notificationEntity.UploadedList,
                    UploadedListName = notificationEntity.UploadedListName,
                    EmailOption = notificationEntity.EmailOption,
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
        public async Task UpdateExpiredNotificationAsync(string notificationId)
        {
            var notificationDataEntity = await this.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                notificationId);

            if (notificationDataEntity != null)
            {
                notificationDataEntity.IsExpiredContentErased = true;

                await this.CreateOrUpdateAsync(notificationDataEntity);
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
