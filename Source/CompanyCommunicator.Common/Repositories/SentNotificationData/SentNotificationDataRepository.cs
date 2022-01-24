// <copyright file="SentNotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
    using Polly;

    /// <summary>
    /// Repository of the notification data in the table storage.
    /// </summary>
    public class SentNotificationDataRepository : BaseRepository<SentNotificationDataEntity>, ISentNotificationDataRepository
    {
        private readonly AdaptiveCardCreator adaptiveCardCreator;
        private readonly string userAppId;
        private readonly BotFrameworkHttpAdapter botAdapter;
        private readonly IAppSettingsService appSettingsService;

        /// <summary>
        /// Initializes a new instance of the <see cref="SentNotificationDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public SentNotificationDataRepository(
            ILogger<SentNotificationDataRepository> logger,
            //IAppSettingsService appSettingsService,
            BotFrameworkHttpAdapter botAdapter,
            IOptions<BotOptions> botOptions,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: SentNotificationDataTableNames.TableName,
                  defaultPartitionKey: SentNotificationDataTableNames.DefaultPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
            this.botAdapter = botAdapter ?? throw new ArgumentNullException(nameof(botAdapter));
            this.userAppId = botOptions?.Value?.UserAppId ?? throw new ArgumentNullException(nameof(botOptions));
            //this.appSettingsService = appSettingsService ?? throw new ArgumentNullException(nameof(appSettingsService));
        }

        /// <inheritdoc/>
        public async Task EnsureSentNotificationDataTableExistsAsync()
        {
            var exists = await this.Table.ExistsAsync();
            if (!exists)
            {
                await this.Table.CreateAsync();
            }
        }

        /// <inheritdoc/>
        public async Task UpdateSentNotificationCardAsync(string notificationId)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0));

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = "No longer available.",
                Size = AdaptiveTextSize.Small,
            });

            //var serviceUrl = await this.appSettingsService.GetServiceUrlAsync();
            var serviceUrl = "https://smba.trafficmanager.net/apac/";

            var sentNotification = await this.GetAllAsync(
                partition: notificationId);

            foreach (var x in sentNotification)
            {
                var y = new SentNotificationDataEntity
                {
                    PartitionKey = x.PartitionKey,
                    ActivityId = x.ActivityId,
                    ConversationId = x.ConversationId,
                };

                if (y.ActivityId != null && y.ConversationId != null)
                {
                    var adaptiveCardAttachment = new Attachment()
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = card,
                    };

                    var activity = MessageFactory.Attachment(adaptiveCardAttachment);
                    activity.Id = y.ActivityId;

                    var conversationReference = new ConversationReference
                    {
                        ServiceUrl = serviceUrl,
                        Conversation = new ConversationAccount
                        {
                            Id = y.ConversationId,
                        },
                    };

                    var turnContext = new TurnContext(this.botAdapter, (Bot.Schema.Activity)activity);

                    int maxNumberOfAttempts = 10;

                    await this.botAdapter.ContinueConversationAsync(
                       botAppId: this.userAppId,
                       reference: conversationReference,
                       callback: async (turnContext, cancellationToken) =>
                       {
                           // Retry it in addition to the original call.
                           var retryPolicy = Policy.Handle<Exception>().WaitAndRetryAsync(maxNumberOfAttempts, p => TimeSpan.FromSeconds(p));
                           await retryPolicy.ExecuteAsync(async () =>
                           {
                               await turnContext.UpdateActivityAsync(activity, cancellationToken);
                           });
                       },
                       cancellationToken: CancellationToken.None);
                }
            }
        }
    }
}