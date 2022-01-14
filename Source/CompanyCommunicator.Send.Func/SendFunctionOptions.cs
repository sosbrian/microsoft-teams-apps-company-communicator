// <copyright file="SendFunctionOptions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func
{
    /// <summary>
    /// Options used to configure the Company Communicator Send Function.
    /// </summary>
    public class SendFunctionOptions
    {
        /// <summary>
        /// Gets or sets the max number of request attempts.
        /// </summary>
        public int MaxNumberOfAttempts { get; set; }

        /// <summary>
        /// Gets or sets the number of seconds to delay before
        /// retrying to send the message.
        /// </summary>
        public double SendRetryDelayNumberOfSeconds { get; set; }

        /// <summary>
        /// Gets or sets the Email Sender for mail Adaptive Card.
        /// </summary>
        public string EmailSenderAadId { get; set; }

        /// <summary>
        /// Gets or sets the TenantId for mail Adaptive Card.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the OriginatorId for Adaptive Card.
        /// </summary>
        public string OriginatorId { get; set; }

        /// <summary>
        /// Gets or sets the AuthorAppId.
        /// </summary>
        public string AuthorAppId { get; set; }

        /// <summary>
        /// Gets or sets the AuthorAppPassword for Adaptive Card.
        /// </summary>
        public string AuthorAppPassword { get; set; }

        /// <summary>
        /// Gets or sets the AppServiceUri for Adaptive Card.
        /// </summary>
        public string AppServiceUri { get; set; }

        /// <summary>
        /// Gets or sets the TaskModuleAppID for Adaptive Card.
        /// </summary>
        public string TaskModuleAppID { get; set; }
    }
}
