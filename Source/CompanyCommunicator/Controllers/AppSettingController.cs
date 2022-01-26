// <copyright file="AppSettingsController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Controllers
{
    using System;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;

    /// <summary>
    /// Controller to get app settings.
    /// </summary>
    [Route("api/settings")]
    [ApiController]
    public class AppSettingsController : ControllerBase
    {
        private readonly BotOptions botOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppSettingsController"/> class.
        /// </summary>
        /// <param name="userAppOptions">User app options.</param>
        public AppSettingsController(
            IOptions<BotOptions> userAppOptions)
        {
            this.botOptions = userAppOptions?.Value ?? throw new ArgumentNullException(nameof(userAppOptions));
        }

        /// <summary>
        /// Get app id and if targeting is enabled.
        /// </summary>
        /// <returns>Required sent notification.</returns>
        [HttpGet]
        public IActionResult GetAppSettings()
        {
            var ConnectionString = this.botOptions.StorageAccountName;
            var SasToken = this.botOptions.SasToken;
            var response = new AppConfigurations()
            {
                StorageAccountName = ConnectionString,
                SasToken = SasToken,
            };

            return this.Ok(response);
        }
    }
}