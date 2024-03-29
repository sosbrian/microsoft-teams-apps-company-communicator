﻿// <copyright file="DraftNotificationSummaryForConsent.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// Draft notification summary (for consent page) model class.
    /// </summary>
    public class DraftNotificationSummaryForConsent
    {
        /// <summary>
        /// Gets or sets Notification Id value.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets Team Names value.
        /// </summary>
        public IEnumerable<string> TeamNames { get; set; }

        /// <summary>
        /// Gets or sets Roster Names value.
        /// </summary>
        public IEnumerable<string> RosterNames { get; set; }

        /// <summary>
        /// Gets or sets Group Names value.
        /// </summary>
        public IEnumerable<string> GroupNames { get; set; }

        /// <summary>
        /// Gets or sets Uploaded List value.
        /// </summary>
        public IEnumerable<string> UploadedList { get; set; }

        /// <summary>
        /// Gets or sets Uploaded List name.
        /// </summary>
        public string UploadedListName { get; set; }

        /// <summary>
        /// Gets or sets email option.
        /// </summary>
        public bool EmailOption { get; set; }

        /// <summary>
        /// Gets or sets Exclusion List value.
        /// </summary>
        public string ExclusionList { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the All Users option is selected.
        /// </summary>
        public bool AllUsers { get; set; }
    }
}
