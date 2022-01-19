// <copyright file="UserDataController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Bot.Schema;

    /// <summary>
    /// Controller for the user data.
    /// </summary>
    [Route("api/user")]
    //[Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class UserController : ControllerBase
    {
        private readonly IUserDataRepository userDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserController"/> class.
        /// </summary>
        /// <param name="userDataRepository">Team data repository instance.</param>
        public UserController(IUserDataRepository userDataRepository)
        {
           this.userDataRepository = userDataRepository ?? throw new ArgumentNullException(nameof(userDataRepository));
        }

        /// <summary>
        /// Get data for all teams.
        /// </summary>
        /// <returns>A list of team data.</returns>
        [HttpGet("{id}")]
        public async Task<IEnumerable<UserDataEntity>> GetAllUsersDataAsync(IEnumerable<string> id) {
            var entities = await this.userDataRepository.GetUserDataEntitiesByIdsAsync(id);
            var result = new List<UserDataEntity>();
            foreach (var entity in entities)
            {
                var user = new UserDataEntity
                {
                    PartitionKey = entity.PartitionKey,
                    RowKey = entity.RowKey,
                    AadId = entity.AadId,
                    UserId = entity.UserId,
                    ConversationId = entity.ConversationId,
                    ServiceUrl = entity.ServiceUrl,
                    TenantId = entity.TenantId,
                    Preference = entity.Preference,
                    UserType = entity.UserType,
                };
                result.Add(user);
            }

            return result;
        }

        // [HttpPost("update")]
        [HttpPut]
        public async Task<IActionResult> UpdateUserPreferenceAsync([FromBody] UserDataEntity userData)
        {
            if (userData == null)
            {
                throw new ArgumentNullException(nameof(userData));
            }

            var userDataEntity = new UserDataEntity
            {
                PartitionKey = userData.PartitionKey,
                RowKey = userData.RowKey,
                AadId = userData.AadId,
                UserId = userData.UserId,
                ConversationId = userData.ConversationId,
                ServiceUrl = userData.ServiceUrl,
                TenantId = userData.TenantId,
                Preference = userData.Preference,
                UserType = userData.UserType,
            };

            //{
            //    PartitionKey = UserDataTableNames.UserDataPartition,
            //    RowKey = userData.AadId,
            //    Preference = userData.Preference,
            //};

            await this.userDataRepository.InsertOrMergeAsync(userDataEntity);
            return this.Ok();
        }
    }
}
