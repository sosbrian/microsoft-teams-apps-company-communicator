// <copyright file="AdaptiveCardCreator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using System;
    using System.Collections.Generic;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive Card Creator service.
    /// </summary>
    public class AdaptiveCardCreator
    {
        private readonly string taskModuleAppID;
        private readonly string appServiceUri;

        /// <summary>
        /// Initializes a new instance of the <see cref="AdaptiveCardCreator"/> class.
        /// </summary>
        /// <param name="options">Bot options.</param>
        public AdaptiveCardCreator(
            IOptions<BotOptions> options)
        {
            if (options is null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            this.taskModuleAppID = options.Value.TaskModuleAppID;
            this.appServiceUri = options.Value.AppServiceUri;
        }

        /// <summary>
        /// Creates an adaptive card.
        /// </summary>
        /// <param name="notificationDataEntity">Notification data entity.</param>
        /// <returns>An adaptive card.</returns>
        public virtual AdaptiveCard CreateAdaptiveCard(NotificationDataEntity notificationDataEntity,
            bool submitted = false)
        {
            return this.CreateAdaptiveCard(
                notificationDataEntity.SenderTemplate,
                notificationDataEntity.Title,
                notificationDataEntity.ImageLink,
                notificationDataEntity.VideoLink,
                notificationDataEntity.Summary,
                notificationDataEntity.Alignment,
                notificationDataEntity.BoldSummary,
                notificationDataEntity.FontSummary,
                notificationDataEntity.FontSizeSummary,
                notificationDataEntity.FontColorSummary,
                notificationDataEntity.Author,
                notificationDataEntity.ButtonTitle,
                notificationDataEntity.ButtonLink,
                notificationDataEntity.ButtonTitle2,
                notificationDataEntity.ButtonLink2,
                notificationDataEntity.ButtonTitle3,
                notificationDataEntity.ButtonLink3,
                notificationDataEntity.ButtonTitle4,
                notificationDataEntity.ButtonLink4,
                notificationDataEntity.ButtonTitle5,
                notificationDataEntity.ButtonLink5,
                notificationDataEntity.SurReaction,
                notificationDataEntity.ReactionQuestion,
                notificationDataEntity.SurFreeText,
                notificationDataEntity.FreeTextQuestion,
                notificationDataEntity.SurYesNo,
                notificationDataEntity.YesNoQuestion,
                notificationDataEntity.SurLinkToSurvey,
                notificationDataEntity.LinkToSurvey,
                notificationDataEntity.Id,
                submitted);
        }

        /// <summary>
        /// Create an adaptive card instance.
        /// </summary>
        /// <param name="senderTemplate">The adaptive card's senderTemplate value.</param>
        /// <param name="title">The adaptive card's title value.</param>
        /// <param name="imageUrl">The adaptive card's image URL.</param>
        /// <param name="videoUrl">The adaptive card's video URL.</param>
        /// <param name="summary">The adaptive card's summary value.</param>
        /// <param name="alignment">The adaptive card's summary alignment value.</param>
        /// <param name="boldSummary">The adaptive card's summary bold value.</param>
        /// <param name="fontSummary">The adaptive card's summary font value.</param>
        /// <param name="fontSizeSummary">The adaptive card's summary font size value.</param>
        /// <param name="fontColorSummary">The adaptive card's summary font color value.</param>
        /// <param name="author">The adaptive card's author value.</param>
        /// <param name="buttonTitle">The adaptive card's button title value.</param>
        /// <param name="buttonUrl">The adaptive card's button url value.</param>
        /// <param name="buttonTitle2">The adaptive card's button title 2 value.</param>
        /// <param name="buttonUrl2">The adaptive card's button url 2 value.</param>
        /// <param name="buttonTitle3">The adaptive card's button title 3 value.</param>
        /// <param name="buttonUrl3">The adaptive card's button url 3 value.</param>
        /// <param name="buttonTitle4">The adaptive card's button title 4 value.</param>
        /// <param name="buttonUrl4">The adaptive card's button url 4 value.</param>
        /// <param name="buttonTitle5">The adaptive card's button title 5 value.</param>
        /// <param name="buttonUrl5">The adaptive card's button url 5 value.</param>
        /// <param name="surReaction">The adaptive card's surReaction value.</param>
        /// <param name="reactionQuestion">The adaptive card's reactionQuestion value.</param>
        /// <param name="surFreeText">The adaptive card's surFreeText value.</param>
        /// <param name="freeTextQuestion">The adaptive card's freeTextQuestion value.</param>
        /// <param name="surYesNo">The adaptive card's surYesNo value.</param>
        /// <param name="yesNoQuestion">The adaptive card's yesNoQuestion value.</param>
        /// <param name="surLinkToSurvey">The adaptive card's surLinkToSurvey value.</param>
        /// <param name="linkToSurvey">The adaptive card's linkToSurvey value.</param>
        /// <returns>The created adaptive card instance.</returns>
        public AdaptiveCard CreateAdaptiveCard(
            string senderTemplate,
            string title,
            string imageUrl,
            string videoUrl,
            string summary,
            string alignment,
            string boldSummary,
            string fontSummary,
            string fontSizeSummary,
            string fontColorSummary,
            string author,
            string buttonTitle,
            string buttonUrl,
            string buttonTitle2,
            string buttonUrl2,
            string buttonTitle3,
            string buttonUrl3,
            string buttonTitle4,
            string buttonUrl4,
            string buttonTitle5,
            string buttonUrl5,
            bool surReaction,
            string reactionQuestion,
            bool surFreeText,
            string freeTextQuestion,
            bool surYesNo,
            string yesNoQuestion,
            bool surLinkToSurvey,
            string linkToSurvey,
            string notificationId,
            bool submitted = false)
        {
            var version = new AdaptiveSchemaVersion(1, 2);
            AdaptiveCard card = new AdaptiveCard(version);
            var tempVideoLink = "https://teams.microsoft.com/l/task/418d0042-3b64-42ed-8d82-cf22461d65ff?url=https://chrischow.ap.ngrok.io/player.html?vid=OhFsua8pjjA&height=700&width=1000&title=YoutubePlayer";
            var summarybold = AdaptiveTextWeight.Default;
            var summaryFontType = AdaptiveFontType.Default;
            var summarySize = AdaptiveTextSize.Default;
            var summaryHorizontalAlignment = AdaptiveHorizontalAlignment.Left;
            var summaryColor = AdaptiveTextColor.Default;
            var encodedSummary = HttpUtility.HtmlEncode(summary);
            var taskmodulevideoURL = "https://teams.microsoft.com/l/task/" + this.taskModuleAppID + "?url=" + this.appServiceUri + "/player.html?vid=" + videoUrl + "&height=700&width=1000&title=Video%20Player";
            if (!string.IsNullOrWhiteSpace(senderTemplate))
            {
                card.Body.Add(new AdaptiveContainer()
                {
                    Bleed = true,
                    BackgroundImage = new AdaptiveBackgroundImage(
                        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAATSURBVChTY7gs6IoHjUpjQYKuAHs0dAUXB9EuAAAAAElFTkSuQmCC",
                        AdaptiveImageFillMode.Repeat,
                        AdaptiveHorizontalAlignment.Center,
                        AdaptiveVerticalAlignment.Center),
                    Items = new List<AdaptiveElement>()
                    {
                        new AdaptiveTextBlock()
                            {
                                Text = senderTemplate,
                                Weight = AdaptiveTextWeight.Bolder,
                                HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                Size = AdaptiveTextSize.Medium,
                                Color = AdaptiveTextColor.Light,
                            },
                    },
                });
            }

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = title,
                Size = AdaptiveTextSize.ExtraLarge,
                Weight = AdaptiveTextWeight.Bolder,
                Wrap = true,
            });

            if (!string.IsNullOrWhiteSpace(imageUrl))
            {
                card.Body.Add(new AdaptiveImage()
                {
                    Url = new Uri(imageUrl, UriKind.RelativeOrAbsolute),
                    Spacing = AdaptiveSpacing.Default,
                    Size = AdaptiveImageSize.Stretch,
                    AltText = string.Empty,
                });
            }

            if (!string.IsNullOrWhiteSpace(summary))
            {
                if (boldSummary == "Bold")
                { summarybold = AdaptiveTextWeight.Bolder; }
                else
                { summarybold = AdaptiveTextWeight.Default; }

                if (fontSummary == "Monospace")
                { summaryFontType = AdaptiveFontType.Monospace; }
                else
                { summaryFontType = AdaptiveFontType.Default; }

                if (fontSizeSummary == "small")
                { summarySize = AdaptiveTextSize.Small; }
                else if (fontSizeSummary == "medium")
                { summarySize = AdaptiveTextSize.Medium; }
                else if (fontSizeSummary == "large")
                { summarySize = AdaptiveTextSize.Large; }
                else if (fontSizeSummary == "extraLarge")
                { summarySize = AdaptiveTextSize.ExtraLarge; }
                else
                { summarySize = AdaptiveTextSize.Default; }

                if (alignment == "center")
                { summaryHorizontalAlignment = AdaptiveHorizontalAlignment.Center; }
                else if (alignment == "right")
                { summaryHorizontalAlignment = AdaptiveHorizontalAlignment.Right; }
                else
                { summaryHorizontalAlignment = AdaptiveHorizontalAlignment.Left; }

                if (fontColorSummary == "dark")
                { summaryColor = AdaptiveTextColor.Dark; }
                else if (fontColorSummary == "light")
                { summaryColor = AdaptiveTextColor.Light; }
                else if (fontColorSummary == "accent")
                { summaryColor = AdaptiveTextColor.Accent; }
                else if (fontColorSummary == "good")
                { summaryColor = AdaptiveTextColor.Good; }
                else if (fontColorSummary == "warning")
                { summaryColor = AdaptiveTextColor.Warning; }
                else if (fontColorSummary == "attention")
                { summaryColor = AdaptiveTextColor.Attention; }
                else
                { summaryColor = AdaptiveTextColor.Default; }

                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = summary,
                    Weight = summarybold,
                    FontType = summaryFontType,
                    Size = summarySize,
                    HorizontalAlignment = summaryHorizontalAlignment,
                    Color = summaryColor,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(author))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = author,
                    Size = AdaptiveTextSize.Small,
                    Weight = AdaptiveTextWeight.Lighter,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(videoUrl))
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = "Watch Video",
                            Url = new Uri(taskmodulevideoURL, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle) && !string.IsNullOrWhiteSpace(buttonUrl))
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = buttonTitle,
                            Url = new Uri(buttonUrl, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle2) && !string.IsNullOrWhiteSpace(buttonUrl2))
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = buttonTitle2,
                            Url = new Uri(buttonUrl2, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle3) && !string.IsNullOrWhiteSpace(buttonUrl3))
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = buttonTitle3,
                            Url = new Uri(buttonUrl3, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle4) && !string.IsNullOrWhiteSpace(buttonUrl4))
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = buttonTitle4,
                            Url = new Uri(buttonUrl4, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle5) && !string.IsNullOrWhiteSpace(buttonUrl5))
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = buttonTitle5,
                            Url = new Uri(buttonUrl5, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            if (surReaction && !submitted)
            {
                var reactchoices = new List<AdaptiveChoice>();
                reactchoices.Add(new AdaptiveChoice()
                {
                    Title = "Extremely satisfied",
                    Value = "1",
                });
                reactchoices.Add(new AdaptiveChoice()
                {
                    Title = "Somewhat satisfied",
                    Value = "2",
                });
                reactchoices.Add(new AdaptiveChoice()
                {
                    Title = "Neither satisfied nor dissatisfied",
                    Value = "3",
                });
                reactchoices.Add(new AdaptiveChoice()
                {
                    Title = "Somewhat dissatisfied",
                    Value = "4",
                });
                reactchoices.Add(new AdaptiveChoice()
                {
                    Title = "Extremely dissatisfied",
                    Value = "5",
                });

                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = reactionQuestion,
                    Wrap = true,
                });

                card.Body.Add(new AdaptiveChoiceSetInput()
                {
                    Id = "Reaction",
                    Style = AdaptiveChoiceInputStyle.Expanded,
                    IsMultiSelect = false,
                    IsRequired = true,
                    Choices = reactchoices,
                });
            }

            if (surFreeText && !submitted)
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = freeTextQuestion,
                    Wrap = true,
                });

                card.Body.Add(new AdaptiveTextInput()
                {
                    Id = "FreeTextSurvey",
                    Placeholder = "Enter Text Here",
                    MaxLength = 500,
                    IsRequired = true,
                    IsMultiline = true,
                });
            }

            if (surYesNo && !submitted)
            {
                var yesnochoices = new List<AdaptiveChoice>();
                yesnochoices.Add(new AdaptiveChoice()
                {
                    Title = "Yes",
                    Value = "Yes",
                });
                yesnochoices.Add(new AdaptiveChoice()
                {
                    Title = "No",
                    Value = "No",
                });

                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = yesNoQuestion,
                    Wrap = true,
                });

                card.Body.Add(new AdaptiveChoiceSetInput()
                {
                    Id = "YesNo",
                    Style = AdaptiveChoiceInputStyle.Expanded,
                    IsMultiSelect = false,
                    IsRequired = true,
                    Choices = yesnochoices,
                });
            }

            if ((surYesNo || surReaction || surFreeText) && !submitted)
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveSubmitAction()
                        {
                            Title = "Submit",
                            Id = "submitSurvey",
                            //Data = "Submit",
                            //Data = JsonConvert.SerializeObject(
                            //    new {msteams = JsonConvert.SerializeObject(
                            //        new {type = "task/fetch"}
                            //        ),
                            //        data = "Invoke"
                            //    }),
                            DataJson = JsonConvert.SerializeObject(
                                new
                                {
                                    notificationId = notificationId,
                                }),
                        },
                    },
                });
            }

            if (submitted)
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = "Submitted",
                    Size = AdaptiveTextSize.Large,
                    Weight = AdaptiveTextWeight.Bolder,
                    HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                    Color = AdaptiveTextColor.Dark,
                    Wrap = true,
                });
            }

            if (submitted)
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = "Thank you for your thoughtful feedback!",
                    Size = AdaptiveTextSize.Medium,
                    Weight = AdaptiveTextWeight.Lighter,
                    HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                    Color = AdaptiveTextColor.Dark,
                    Wrap = true,
                });
            }

            if (surLinkToSurvey)
            {
                card.Body.Add(new AdaptiveActionSet()
                {
                    Actions = new List<AdaptiveAction>()
                    {
                        new AdaptiveOpenUrlAction()
                        {
                            Title = "Open Survey",
                            Url = new Uri(linkToSurvey, UriKind.RelativeOrAbsolute),
                        },
                    },
                });
            }

            return card;
        }
    }
}
