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
        /// Gets or sets PriLanguage value.
        /// </summary>
        public string PriLanguage { get; set; }

        /// <summary>
        /// Gets or sets SecLanguage value.
        /// </summary>
        public string SecLanguage { get; set; }

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
        /// Gets or sets SecSenderTemplate value.
        /// </summary>
        public string SecSenderTemplate { get; set; }

        /// <summary>
        /// Gets or sets SecTitle value.
        /// </summary>
        public string SecTitle { get; set; }

        /// <summary>
        /// Gets or sets the SecImage Link value.
        /// </summary>
        public string SecImageLink { get; set; }

        /// <summary>
        /// Gets or sets the SecVideo Link value.
        /// </summary>
        public string SecVideoLink { get; set; }

        /// <summary>
        /// Gets or sets the SecSummary value.
        /// </summary>
        public string SecSummary { get; set; }

        /// <summary>
        /// Gets or sets the Secsummary alignment of the notification's content.
        /// </summary>
        public string SecAlignment { get; set; }

        /// <summary>
        /// Gets or sets the Secsummary bold of the notification's content.
        /// </summary>
        public string SecBoldSummary { get; set; }

        /// <summary>
        /// Gets or sets the Secsummary font of the notification's content.
        /// </summary>
        public string SecFontSummary { get; set; }

        /// <summary>
        /// Gets or sets the Secsummary font size of the notification's content.
        /// </summary>
        public string SecFontSizeSummary { get; set; }

        /// <summary>
        /// Gets or sets the Secsummary font Color of the notification's content.
        /// </summary>
        public string SecFontColorSummary { get; set; }

        /// <summary>
        /// Gets or sets the SecAuthor value.
        /// </summary>
        public string SecAuthor { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Title value.
        /// </summary>
        public string SecButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Link value.
        /// </summary>
        public string SecButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Title value.
        /// </summary>
        public string SecButtonTitle2 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Link value.
        /// </summary>
        public string SecButtonLink2 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Title value.
        /// </summary>
        public string SecButtonTitle3 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Link value.
        /// </summary>
        public string SecButtonLink3 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Title value.
        /// </summary>
        public string SecButtonTitle4 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Link value.
        /// </summary>
        public string SecButtonLink4 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Title value.
        /// </summary>
        public string SecButtonTitle5 { get; set; }

        /// <summary>
        /// Gets or sets the SecButton Link value.
        /// </summary>
        public string SecButtonLink5 { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey Reaction value.
        /// </summary>
        public bool SecSurReaction { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey Reaction question value.
        /// </summary>
        public string SecReactionQuestion { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey Free Text value.
        /// </summary>
        public bool SecSurFreeText { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey Free Text question value.
        /// </summary>
        public string SecFreeTextQuestion { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey yes/no value.
        /// </summary>
        public bool SecSurYesNo { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey yes/no question value.
        /// </summary>
        public string SecYesNoQuestion { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey Link value.
        /// </summary>
        public bool SecSurLinkToSurvey { get; set; }

        /// <summary>
        /// Gets or sets the SecSurvey Link link value.
        /// </summary>
        public string SecLinkToSurvey { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }

        /// <summary>	
        /// Gets or sets the DateTime the notification's was scheduled to be sent.	
        /// </summary>	
        public DateTime? ScheduledDate { get; set; }
        /// <summary>	
        /// Gets or sets a value indicating whether the expiry date is set.	
        /// </summary>	
        public bool IsExpirySet { get; set; }
        /// <summary>	
        /// Gets or sets the Expiry DateTime the notification if it is was scheduled to be sent.	
        /// </summary>	
        public DateTime? ExpiryDate { get; set; }
        /// <summary>	
        /// Gets or sets a value indicating whether the expired content is erased.	
        /// </summary>	
        public bool IsExpiredContentErased { get; set; }

        /// <summary>	
        /// Gets or sets a value indicating whether the notification is scheduled.	
        /// </summary>	
        public bool IsScheduled { get; set; }
    }
}
