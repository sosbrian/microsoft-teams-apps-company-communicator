// <copyright file="BaseNotification.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;

    /// <summary>
    /// Base notification model class.
    /// </summary>
    public class BaseNotification
    {
        /// <summary>
        /// Gets or sets Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Template value.
        /// </summary>
        public string Template { get; set; }

        /// <summary>
        /// Gets or sets SenderTemplate value.
        /// </summary>
        public string SenderTemplate { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Image Link value.
        /// </summary>
        public string ImageLink { get; set; }

        /// <summary>
        /// Gets or sets the Video Link value.
        /// </summary>
        public string VideoLink { get; set; }

        /// <summary>
        /// Gets or sets the Summary value.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the summary alignment of the notification's content.
        /// </summary>
        public string Alignment { get; set; }

        /// <summary>
        /// Gets or sets the summary bold of the notification's content.
        /// </summary>
        public string BoldSummary { get; set; }

        /// <summary>
        /// Gets or sets the summary font of the notification's content.
        /// </summary>
        public string FontSummary { get; set; }

        /// <summary>
        /// Gets or sets the summary font size of the notification's content.
        /// </summary>
        public string FontSizeSummary { get; set; }

        /// <summary>
        /// Gets or sets the summary font Color of the notification's content.
        /// </summary>
        public string FontColorSummary { get; set; }

        /// <summary>
        /// Gets or sets the Author value.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle2 { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink2 { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle3 { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink3 { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle4 { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink4 { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle5 { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink5 { get; set; }

        /// <summary>
        /// Gets or sets the Survey Reaction value.
        /// </summary>
        public bool SurReaction { get; set; }

        /// <summary>
        /// Gets or sets the Survey Reaction question value.
        /// </summary>
        public string ReactionQuestion { get; set; }

        /// <summary>
        /// Gets or sets the Survey Free Text value.
        /// </summary>
        public bool SurFreeText { get; set; }

        /// <summary>
        /// Gets or sets the Survey Free Text question value.
        /// </summary>
        public string FreeTextQuestion { get; set; }

        /// <summary>
        /// Gets or sets the Survey yes/no value.
        /// </summary>
        public bool SurYesNo { get; set; }

        /// <summary>
        /// Gets or sets the Survey yes/no question value.
        /// </summary>
        public string YesNoQuestion { get; set; }

        /// <summary>
        /// Gets or sets the Survey Link value.
        /// </summary>
        public bool SurLinkToSurvey { get; set; }

        /// <summary>
        /// Gets or sets the Survey Link link value.
        /// </summary>
        public string LinkToSurvey { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }
    }
}
