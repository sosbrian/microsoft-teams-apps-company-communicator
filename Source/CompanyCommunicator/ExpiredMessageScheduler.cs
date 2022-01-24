// <copyright file="SendMessageScheduler.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToSendQueue;

    /// <summary>
    /// Register background timed service to erase content of expired messages.
    /// </summary>
    public class ExpiredMessageScheduler : IHostedService, IDisposable
    {
        private readonly ILogger<SendMessageScheduler> smslogger;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly IPrepareToSendQueue prepareToSendQueue;
        private readonly IDataQueue dataQueue;
        private readonly double forceCompleteMessageDelayInSeconds;
        private Timer smstimer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpiredMessageScheduler"/> class.
        /// </summary>
        /// <param name="logger">system logger.</param>
        /// <param name="factory">factory.</param>
        public ExpiredMessageScheduler(ILogger<SendMessageScheduler> logger, IServiceScopeFactory factory)
        {
            this.smslogger = logger;
            this.notificationDataRepository = factory.CreateScope().ServiceProvider.GetRequiredService<INotificationDataRepository>();
            this.sentNotificationDataRepository = factory.CreateScope().ServiceProvider.GetRequiredService<ISentNotificationDataRepository>();
            this.prepareToSendQueue = factory.CreateScope().ServiceProvider.GetRequiredService<IPrepareToSendQueue>();
            this.dataQueue = factory.CreateScope().ServiceProvider.GetRequiredService<IDataQueue>();
            this.forceCompleteMessageDelayInSeconds = 86400;
        }

        /// <summary>
        /// Start the service <see cref="StartAsync"/>.
        /// </summary>
        /// <param name="stoppingToken">system logger.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public Task StartAsync(CancellationToken stoppingToken)
        {
            this.smslogger.LogInformation("[CC Expiry Scheduler] Hosted Service is running.");

            this.smstimer = new Timer(this.DoWork, null, TimeSpan.Zero, TimeSpan.FromMinutes(5));

            return Task.CompletedTask;
        }

        /// <summary>
        /// Stops the service <see cref="StopAsync"/>.
        /// </summary>
        /// <param name="stoppingToken">This is the cancellation token.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public Task StopAsync(CancellationToken stoppingToken)
        {
            this.smslogger.LogInformation("[CC Expiry Scheduler] Hosted Service is stopping.");

            this.smstimer?.Change(Timeout.Infinite, 0);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Disposes the service.
        /// </summary>
        public void Dispose()
        {
            this.smstimer?.Dispose();
        }

        private async void DoWork(object state)
        {
            DateTime now = DateTime.Now;

            this.smslogger.LogInformation(
                "[CC Expiry Scheduler] is processing expired messages before {Now}.", now);

            try
            {
                var notificationEntities = await this.notificationDataRepository.GetNonErasedExpiredNotificationsAsync();
                foreach (var notificationEntity in notificationEntities)
                {
                    this.smslogger.LogInformation("[CC Expiry Scheduler] sending notification: {0}", notificationEntity.Title);
                    await this.notificationDataRepository.UpdateExpiredNotificationAsync(notificationEntity.Id);
                    await this.sentNotificationDataRepository.UpdateSentNotificationCardAsync(notificationEntity.Id);
                }
            }
            catch (Exception ex)
            {
                this.smslogger.LogError(ex.ToString());
            }
        }
    }
}