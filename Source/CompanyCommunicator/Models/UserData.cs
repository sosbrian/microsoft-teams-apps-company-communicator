// <copyright file="TeamData.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    /// <summary>
    /// Teams data model class.
    /// </summary>
    public class UserData
    {
        /// <summary>
        /// Gets or sets team Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets name.
        /// </summary>
        public string Name { get; set; }

        public string AadId { get; set; }

        public string Preference { get; set; }
    }
}
