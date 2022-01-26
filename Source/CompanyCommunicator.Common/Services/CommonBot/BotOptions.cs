// <copyright file="BotOptions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot
{
    /// <summary>
    /// Options used for holding metadata for the bot.
    /// </summary>
    public class BotOptions
    {
        /// <summary>
        /// Gets or sets the Microsoft app ID for the user bot.
        /// </summary>
        public string UserAppId { get; set; }

        /// <summary>
        /// Gets or sets the Microsoft app password for the user bot.
        /// </summary>
        public string UserAppPassword { get; set; }

        /// <summary>
        /// Gets or sets the Microsoft app ID for the author bot.
        /// </summary>
        public string AuthorAppId { get; set; }

        /// <summary>
        /// Gets or sets the Microsoft app password for the author bot.
        /// </summary>
        public string AuthorAppPassword { get; set; }

        /// <summary>
        /// Gets or sets the Cient App ID for Task Module.
        /// </summary>
        public string TaskModuleAppID { get; set; }

        /// <summary>
        /// Gets or sets the App Service Uri for Task Module.
        /// </summary>
        public string AppServiceUri { get; set; }

        /// <summary>
        /// Gets or sets the StorageAccountName.
        /// </summary>
        public string StorageAccountName { get; set; }

        /// <summary>
        /// Gets or sets the StorageAccountName.
        /// </summary>
        public string SasToken { get; set; }
    }
}
