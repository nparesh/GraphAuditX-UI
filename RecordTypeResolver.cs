// ---------------------------------------------------------------------------
// <copyright file="RecordTypeResolver.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Office.Audit.Schema
{
    using System;

    using Microsoft.Office.Audit.Schema.AeD;
    using Microsoft.Office.Audit.Schema.AzureActiveDirectory;
    using Microsoft.Office.Audit.Schema.Compliance.DLP;
    using Microsoft.Office.Audit.Schema.CRM;
    using Microsoft.Office.Audit.Schema.DataCenterSecurity;
    using Microsoft.Office.Audit.Schema.Discovery;
    using Microsoft.Office.Audit.Schema.ExchangeAdmin;
    using Microsoft.Office.Audit.Schema.ExchangeMailbox;
    using Microsoft.Office.Audit.Schema.MailSubmission;
    using Microsoft.Office.Audit.Schema.MicrosoftStream;
    using Microsoft.Office.Audit.Schema.MicrosoftTeams;
    using Microsoft.Office.Audit.Schema.Monitoring;
    using Microsoft.Office.Audit.Schema.OneDrive;
    using Microsoft.Office.Audit.Schema.PowerBI;
    using Microsoft.Office.Audit.Schema.SecurityComplianceCenter;
    using Microsoft.Office.Audit.Schema.SharePoint;
    using Microsoft.Office.Audit.Schema.SkypeForBusiness;
    using Microsoft.Office.Audit.Schema.Sway;
    using Microsoft.Office.Audit.Schema.ThreatFinder;
    using Microsoft.Office.Audit.Schema.ThreatIntelligence;
    using Microsoft.Office.Audit.Schema.Yammer;
    
    /// <summary>
    /// Implements IRecordTypeResolver for non-SLAPI records
    /// </summary>
    public class DefaultRecordTypeResolver : IRecordTypeResolver
    {
        /// <summary>
        /// Singleton instance
        /// </summary>
        private static IRecordTypeResolver instance = new DefaultRecordTypeResolver();

        /// <summary>
        /// Singleton instance
        /// </summary>
        public static IRecordTypeResolver Instance 
        {
            get { return instance; }
        }

        /// <summary>
        /// Gets the type for a record type.
        /// </summary>
        /// <param name="recordType">The record type.</param>
        /// <returns>The type of the record object.</returns>
        public Type GetAuditRecordType(AuditLogRecordType recordType)
        {
            switch (recordType)
            {
                case AuditLogRecordType.AzureActiveDirectory:
                    return typeof(AzureActiveDirectoryAuditRecord);

                case AuditLogRecordType.AzureActiveDirectoryAccountLogon:
                    return typeof(AzureActiveDirectoryAccountLogonAuditRecord);

                case AuditLogRecordType.CRM:
                    return typeof(CRMBaseAuditRecord);

                case AuditLogRecordType.ExchangeAdmin: 
                    return typeof(ExchangeAdminAuditRecord);

                case AuditLogRecordType.ExchangeItem: 
                    return typeof(ExchangeMailboxAuditRecord);

                case AuditLogRecordType.ExchangeItemGroup: 
                    return typeof(ExchangeMailboxAuditGroupRecord);

                case AuditLogRecordType.ExchangeAggregatedOperation:
                    return typeof(ExchangeAggregatedOperationRecord);

                case AuditLogRecordType.DataCenterSecurityCmdlet:
                    return typeof(DataCenterSecurityCmdletAuditRecord);

                case AuditLogRecordType.SecurityComplianceCenterEOPCmdlet:
                    return typeof(SecurityComplianceCenterEOPCmdletAuditRecord);

                case AuditLogRecordType.SharePoint: 
                    return typeof(SharePointAuditRecord);

                case AuditLogRecordType.SyntheticProbe: 
                    return typeof(SyntheticProbeAuditRecord);

                case AuditLogRecordType.SharePointFileOperation: 
                    return typeof(SharePointFileOperationAuditRecord);

                case AuditLogRecordType.SharePointSharingOperation: 
                    return typeof(SharePointSharingOperationAuditRecord);

                case AuditLogRecordType.SharePointListOperation:
                    return typeof(SharePointListOperationAuditRecord);

                case AuditLogRecordType.Project:
                    return typeof(ProjectAuditRecord);

                case AuditLogRecordType.OneDrive: 
                     return typeof(OneDriveAuditRecord);

                case AuditLogRecordType.ComplianceDLPSharePoint:
                     return typeof(ComplianceDLPSharePointAuditRecord);

                case AuditLogRecordType.ComplianceDLPExchange:
                     return typeof(ComplianceDLPExchangeAuditRecord);

                case AuditLogRecordType.Sway:
                     return typeof(SwayAuditRecord);

                case AuditLogRecordType.SkypeForBusinessUsersBlocked:
                     return typeof(SkypeForBusinessUsersBlockedAuditRecord);

                case AuditLogRecordType.AzureActiveDirectoryStsLogon:
                     return typeof(AzureActiveDirectoryStsLogonAuditRecord);

                case AuditLogRecordType.SkypeForBusinessPSTNUsage:
                     return typeof(SkypeForBusinessPSTNUsageAuditRecord);

                case AuditLogRecordType.SkypeForBusinessCmdlets:
                     return typeof(SkypeForBusinessCmdletsAuditRecord);

                case AuditLogRecordType.PowerBIAudit:
                    return typeof(PowerBIAuditRecord);

                case AuditLogRecordType.Yammer:
                    return typeof(YammerAuditRecord);

                case AuditLogRecordType.Discovery:
                    return typeof(DiscoveryAuditRecord);

                case AuditLogRecordType.MicrosoftStream:
                    return typeof(MicrosoftStreamAuditRecord);

                case AuditLogRecordType.MicrosoftTeams:
                    return typeof(MicrosoftTeamsAuditRecord);

                case AuditLogRecordType.MicrosoftTeamsAddOns:
                    return typeof(MicrosoftTeamsAddOnsAuditRecord);

                case AuditLogRecordType.MicrosoftTeamsSettingsOperation:
                    return typeof(MicrosoftTeamsSettingsOperationAuditRecord);

                case AuditLogRecordType.ThreatFinder:
                    return typeof(ThreatFinderAuditRecord);

                case AuditLogRecordType.ThreatIntelligence:
                    return typeof(ThreatIntelligenceMailData);

                case AuditLogRecordType.MailSubmission:
                    return typeof(MailSubmissionData);

                case AuditLogRecordType.AeD:
                    return typeof(AeDAuditRecord);

                default:
                    throw new ArgumentException("Invalid AuditLogRecordType: " + recordType, "recordType");
            }
        }
    }
}
