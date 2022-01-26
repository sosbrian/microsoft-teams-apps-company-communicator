// <copyright file="UserDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the user data stored in the table storage.
    /// </summary>
    public class UserDataRepository : BaseRepository<UserDataEntity>, IUserDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public UserDataRepository(
            ILogger<UserDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: UserDataTableNames.TableName,
                  defaultPartitionKey: UserDataTableNames.UserDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task<string> GetDeltaLinkAsync()
        {
            try
            {
                var operation = TableOperation.Retrieve<UsersSyncEntity>(UserDataTableNames.UsersSyncDataPartition, UserDataTableNames.AllUsersDeltaLinkRowKey);
                var result = await this.Table.ExecuteAsync(operation);
                var entity = result.Result as UsersSyncEntity;
                return entity?.Value;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task SetDeltaLinkAsync(string deltaLink)
        {
            if (string.IsNullOrEmpty(deltaLink))
            {
                throw new ArgumentNullException(nameof(deltaLink));
            }

            var entity = new UsersSyncEntity()
            {
                PartitionKey = UserDataTableNames.UsersSyncDataPartition,
                RowKey = UserDataTableNames.AllUsersDeltaLinkRowKey,
                Value = deltaLink,
            };

            try
            {
                var operation = TableOperation.InsertOrReplace(entity);
                await this.Table.ExecuteAsync(operation);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        // public async Task<IEnumerable<string>> GetUserPreferenceByIdsAsync(IEnumerable<string> ids)
        public async Task<IEnumerable<UserDataEntity>> GetSortedUserAsync()
        {
            var userDataEntities = await this.GetAllAsync();
            var sortedSet = new SortedSet<UserDataEntity>(userDataEntities, new UserDataEntityComparer());
            return sortedSet;
            // if (ids == null || !ids.Any())
            // {
            //     return new List<string>();
            // }

            // var rowKeysFilter = this.GetRowKeysFilter(ids);
            // var userDataEntities = await this.GetWithFilterAsync(rowKeysFilter);

            // return userDataEntities.Select(p => p.Name).OrderBy(p => p);
        }

        private class UserDataEntityComparer : IComparer<UserDataEntity>
        {
            public int Compare(UserDataEntity x, UserDataEntity y)
            {
                return x.AadId.CompareTo(y.AadId);
            }
        }

        /// <inheritdoc/>
        public async Task<UserDataEntity> GetUserDataEntitiesByIdsAsync(string userId)
        {
            //var rowKeysFilter = this.GetRowKeysFilter(userIds);

            //return await this.GetWithFilterAsync(rowKeysFilter);
            //var userDataEntity = await this.GetAsync(
            //    UserDataTableNames.UserDataPartition,
            //    userId);
            return await this.GetAsync(
                UserDataTableNames.UserDataPartition,
                userId);
        }
    }
}
