// <copyright file="SyncGroupMembersActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.DurableTask;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.User;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Extensions;

    /// <summary>
    /// Syncs group members to Sent notification table.
    /// </summary>
    public class SyncUploadMembersActivity
    {
        private readonly IUsersService usersService;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly IUserDataRepository userDataRepository;
        private readonly IUserTypeService userTypeService;
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncUploadMembersActivity"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="notificationDataRepository">Notifications data repository.</param>
        /// <param name="usersService">Group members service.</param>
        /// <param name="userDataRepository">User Data repository.</param>
        /// <param name="userTypeService">User Type service.</param>
        /// <param name="localizer">Localization service.</param>
        public SyncUploadMembersActivity(
            ISentNotificationDataRepository sentNotificationDataRepository,
            INotificationDataRepository notificationDataRepository,
            IUsersService usersService,
            IUserDataRepository userDataRepository,
            IUserTypeService userTypeService,
            IStringLocalizer<Strings> localizer)
        {
            this.usersService = usersService ?? throw new ArgumentNullException(nameof(usersService));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.userDataRepository = userDataRepository ?? throw new ArgumentNullException(nameof(userDataRepository));
            this.userTypeService = userTypeService ?? throw new ArgumentNullException(nameof(userTypeService));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
        }

        /// <summary>
        /// Syncs group members to Sent notification table.
        /// </summary>
        /// <param name="input">Input.</param>
        /// <param name="log">Logging service.</param>
        /// <returns>It returns the group transitive members first page and next page url.</returns>
        [FunctionName(FunctionNames.SyncUploadMembersActivity)]
        public async Task RunAsync(
        [ActivityTrigger] (string notificationId, IEnumerable<string> uploadL, string exclusionL) input, ILogger log)
        {
            if (input.notificationId == null)
            {
                throw new ArgumentNullException(nameof(input.notificationId));
            }

            if (input.uploadL == null)
            {
                throw new ArgumentNullException(nameof(input.uploadL));
            }

            if (log == null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            var notificationId = input.notificationId;
            var uploadL = input.uploadL;
            var exclusionL = input.exclusionL;
            //string[] string2Array = exclusionL.Split(";");
            //string[] string2Array = { "chris.chow@soshk.com" };

            try
            {
                // Get all members.
                var users = await this.usersService.GetUserEmailAsync(uploadL);

                // Convert to Recipients
                //var recipients = await this.GetRecipientsAsync(notificationId, users, string2Array);
                var recipients = await this.GetRecipientsAsync(notificationId, users, exclusionL);

                if (!recipients.IsNullOrEmpty())
                {
                    // Store.
                    await this.sentNotificationDataRepository.BatchInsertOrMergeAsync(recipients);
                }
            }
            catch (Exception ex)
            {
                var errorMessage = this.localizer.GetString("FailedToGetMembersForGroupFormat", uploadL, ex.Message);
                log.LogError(ex, errorMessage);
                await this.notificationDataRepository.SaveWarningInNotificationDataEntityAsync(notificationId, errorMessage);
            }
        }

        /// <summary>
        /// Reads corresponding user entity from User table and creates a recipient for every user.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <param name="users">Users.</param>
        /// <param name="arrExclusion">Exclusion List Array.</param>
        /// <returns>List of recipients.</returns>
        private async Task<IEnumerable<SentNotificationDataEntity>> GetRecipientsAsync(string notificationId, IEnumerable<User> users, string arrExclusion)
        {
            var recipients = new ConcurrentBag<SentNotificationDataEntity>();
            string[] arrString = arrExclusion.Split(";");
            // Get User Entities.
            var maxParallelism = Math.Min(100, users.Count());
            await Task.WhenAll(users.ForEachAsync(maxParallelism, async user =>
            {
                var userEntity = await this.userDataRepository.GetAsync(UserDataTableNames.UserDataPartition, user.Id);

                // This is to set the type of user(exisiting only, new ones will be skipped) to identify later if it is member or guest.
                var userType = user.UserPrincipalName.GetUserType();
                if (userEntity == null && userType.Equals(UserType.Guest, StringComparison.OrdinalIgnoreCase))
                {
                    // Skip processing new Guest users.
                    return;
                }

                //if(stringArray.Any(stringToCheck.Contains))
                //if (string.Equals(user.UserPrincipalName, exclusionL, StringComparison.OrdinalIgnoreCase))
                //if (arrExclusion.All(user.UserPrincipalName.Contains))
                //if (arrExclusion.Contains(user.UserPrincipalName))
                if (arrString.Contains(user.UserPrincipalName, StringComparer.OrdinalIgnoreCase))
                {
                    return;
                }

                await this.userTypeService.UpdateUserTypeForExistingUserAsync(userEntity, userType);
                if (userEntity == null)
                {
                    userEntity = new UserDataEntity()
                    {
                        AadId = user.Id,
                        UserType = userType,
                        Preference = userEntity.Preference,
                    };
                }

                recipients.Add(userEntity.CreateInitialSentNotificationDataEntity(partitionKey: notificationId));
            }));

            return recipients;
        }
    }
}
