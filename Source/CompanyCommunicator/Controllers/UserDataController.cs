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
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the user data.
    /// </summary>
    [Route("api/userD")]
    //[Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class UserDataController : ControllerBase
    {
        private readonly IUsersService usersService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserDataController"/> class.
        /// </summary>
        /// <param name="usersService">Team data repository instance.</param>
        public UserDataController(IUsersService usersService)
        {
            this.usersService = usersService ?? throw new ArgumentNullException(nameof(usersService));
        }

        /// <summary>
        /// Get data for all teams.
        /// </summary>
        /// <returns>A list of team data.</returns>
        [HttpGet]
        public async Task<IEnumerable<UserData>> GetAllUserDataAsync()
        {
            var tuple = await this.usersService.GetAllUsersAsync();
            var users = new List<UserData>();
            foreach (var entity in tuple.Item1)
            {
                var user = new UserData
                {
                    Id = entity.Id,
                    Name = entity.DisplayName,
                };
                users.Add(user);
            }

            return users;
        }
    }
}
