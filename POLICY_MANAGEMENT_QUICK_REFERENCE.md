# Microsoft 365 Policy Management - Quick Reference Guide

## Available Policy Management Tools

### 1. Retention Policies
**Tool**: `manage_retention_policies`

**Actions**: list, get, create, update, delete

**Example - Create 90-day retention policy**:
```json
{
  "action": "create",
  "displayName": "90 Day Email Retention",
  "description": "Retain emails for 90 days then delete",
  "isEnabled": true,
  "retentionSettings": {
    "retentionDuration": 90,
    "retentionAction": "KeepAndDelete",
    "deletionType": "AfterRetentionPeriod"
  },
  "locations": {
    "exchangeEmail": true,
    "teamsChats": true
  }
}
```

### 2. Sensitivity Labels
**Tool**: `manage_sensitivity_labels`

**Actions**: list, get, create, update, delete, publish

**Example - Create confidential label with encryption**:
```json
{
  "action": "create",
  "displayName": "Confidential",
  "description": "Confidential data with encryption",
  "tooltip": "Use for sensitive internal data",
  "priority": 10,
  "settings": {
    "encryption": {
      "enabled": true
    },
    "contentMarking": {
      "watermarkText": "CONFIDENTIAL",
      "headerText": "Confidential - Internal Use Only"
    }
  }
}
```

### 3. Information Protection Policies
**Tool**: `manage_information_protection_policies`

**Actions**: list, get, create, update, delete

**Example - Create mandatory labeling policy**:
```json
{
  "action": "create",
  "displayName": "Mandatory Labeling Policy",
  "description": "Require labels on all documents",
  "settings": {
    "mandatoryLabelPolicy": true,
    "requireJustification": true
  }
}
```

### 4. Conditional Access Policies
**Tool**: `manage_conditional_access_policies`

**Actions**: list, get, create, update, delete, enable, disable

**Example - Require MFA for external access**:
```json
{
  "action": "create",
  "displayName": "MFA for External Access",
  "state": "enabled",
  "conditions": {
    "users": {
      "includeUsers": ["All"]
    },
    "applications": {
      "includeApplications": ["All"]
    },
    "locations": {
      "includeLocations": ["All"],
      "excludeLocations": ["AllTrusted"]
    }
  },
  "grantControls": {
    "operator": "OR",
    "builtInControls": ["mfa"]
  }
}
```

**Example - Require compliant device for admin roles**:
```json
{
  "action": "create",
  "displayName": "Compliant Device for Admins",
  "state": "enabled",
  "conditions": {
    "users": {
      "includeRoles": ["62e90394-69f5-4237-9190-012177145e10"]
    },
    "applications": {
      "includeApplications": ["All"]
    }
  },
  "grantControls": {
    "operator": "AND",
    "builtInControls": ["mfa", "compliantDevice"]
  }
}
```

### 5. Defender for Office 365 Policies
**Tool**: `manage_defender_policies`

**Actions**: list, get, create, update, delete

**Policy Types**: safeAttachments, safeLinks, antiPhishing, antiMalware, antiSpam

**Example - Create Safe Attachments policy**:
```json
{
  "action": "create",
  "policyType": "safeAttachments",
  "displayName": "Block Malicious Attachments",
  "settings": {
    "action": "Block",
    "actionOnError": true
  },
  "appliedTo": {
    "recipientDomains": ["contoso.com"]
  }
}
```

**Example - Create Safe Links policy**:
```json
{
  "action": "create",
  "policyType": "safeLinks",
  "displayName": "Scan All URLs",
  "settings": {
    "scanUrls": true,
    "enableForInternalSenders": true,
    "trackClicks": true,
    "allowClickThrough": false
  },
  "appliedTo": {
    "recipientDomains": ["contoso.com"]
  }
}
```

**Example - Create Anti-Phishing policy**:
```json
{
  "action": "create",
  "policyType": "antiPhishing",
  "displayName": "Advanced Anti-Phishing",
  "settings": {
    "enableMailboxIntelligence": true,
    "enableSpoofIntelligence": true,
    "enableUnauthenticatedSender": true
  }
}
```

### 6. Microsoft Teams Policies
**Tool**: `manage_teams_policies`

**Actions**: list, get, create, update, delete, assign

**Policy Types**: messaging, meeting, calling, appSetup, updateManagement

**Example - Create meeting policy with recording**:
```json
{
  "action": "create",
  "policyType": "meeting",
  "displayName": "Standard Meeting Policy",
  "settings": {
    "allowMeetNow": true,
    "allowIPVideo": true,
    "allowCloudRecording": true,
    "allowTranscription": true,
    "allowWhiteboard": true,
    "allowSharedNotes": true
  }
}
```

**Example - Create messaging policy**:
```json
{
  "action": "create",
  "policyType": "messaging",
  "displayName": "Restricted Messaging",
  "settings": {
    "allowUserChat": true,
    "allowUserEditMessage": true,
    "allowUserDeleteMessage": false,
    "allowGiphy": true,
    "giphyRatingType": "Strict",
    "allowMemes": false,
    "allowStickers": true
  }
}
```

**Example - Assign policy to users**:
```json
{
  "action": "assign",
  "policyType": "meeting",
  "policyId": "policy-id-here",
  "assignTo": {
    "users": ["user1@contoso.com", "user2@contoso.com"],
    "groups": ["group-id-1", "group-id-2"]
  }
}
```

### 7. Exchange Online Policies
**Tool**: `manage_exchange_policies`

**Actions**: list, get, create, update, delete

**Policy Types**: addressBook, outlookWebApp, activeSyncMailbox, retentionPolicy, dlpPolicy

**Example - Create ActiveSync policy**:
```json
{
  "action": "create",
  "policyType": "activeSyncMailbox",
  "displayName": "Secure Mobile Devices",
  "settings": {
    "devicePasswordEnabled": true,
    "alphanumericDevicePasswordRequired": true,
    "minDevicePasswordLength": 8,
    "maxDevicePasswordFailedAttempts": 5,
    "maxInactivityTimeDeviceLock": 15,
    "deviceEncryptionEnabled": true,
    "requireDeviceEncryption": true,
    "allowCamera": false,
    "allowWiFi": true,
    "allowBrowser": true
  }
}
```

**Example - Create Outlook Web App policy**:
```json
{
  "action": "create",
  "policyType": "outlookWebApp",
  "displayName": "Standard OWA Policy",
  "settings": {
    "calendarEnabled": true,
    "contactsEnabled": true,
    "tasksEnabled": true,
    "journalEnabled": false,
    "notesEnabled": true,
    "remindersAndNotificationsEnabled": true,
    "premiumClientEnabled": true,
    "rulesEnabled": true,
    "publicFoldersEnabled": false,
    "changePasswordEnabled": true
  }
}
```

### 8. SharePoint Governance Policies
**Tool**: `manage_sharepoint_governance_policies`

**Actions**: list, get, create, update, delete

**Policy Types**: sharingPolicy, accessPolicy, informationBarrier, retentionLabel

**Example - Create sharing policy**:
```json
{
  "action": "create",
  "policyType": "sharingPolicy",
  "displayName": "Restricted External Sharing",
  "settings": {
    "sharingCapability": "ExternalUserSharingOnly",
    "requireAcceptanceForExternalUsers": true,
    "requireAnonymousLinksExpireInDays": 30,
    "defaultSharingLinkType": "Internal",
    "preventExternalUsersFromResharing": true
  }
}
```

**Example - Create access policy with conditional access**:
```json
{
  "action": "create",
  "policyType": "accessPolicy",
  "displayName": "Limited Access for Unmanaged Devices",
  "settings": {
    "conditionalAccessPolicy": "AllowLimitedAccess",
    "limitedAccessFileType": "OfficeOnlineFilesOnly",
    "allowDownload": false,
    "allowPrint": false,
    "allowCopy": false
  }
}
```

### 9. Security and Compliance Alert Policies
**Tool**: `manage_security_alert_policies`

**Actions**: list, get, create, update, delete, enable, disable

**Example - Create DLP alert policy**:
```json
{
  "action": "create",
  "displayName": "High Severity DLP Violations",
  "category": "DataLossPrevention",
  "severity": "High",
  "isEnabled": true,
  "conditions": {
    "activityType": "DlpRuleMatch",
    "userType": "Regular"
  },
  "actions": {
    "notifyUsers": ["admin@contoso.com", "security@contoso.com"],
    "escalateToAdmin": true,
    "threshold": {
      "value": 5,
      "timeWindow": 60
    }
  }
}
```

**Example - Create threat management alert**:
```json
{
  "action": "create",
  "displayName": "Suspicious Sign-In Activity",
  "category": "ThreatManagement",
  "severity": "High",
  "conditions": {
    "activityType": "SuspiciousSignIn",
    "userType": "Admin"
  },
  "actions": {
    "notifyUsers": ["security-team@contoso.com"],
    "escalateToAdmin": true
  }
}
```

## Common Workflows

### Complete DLP Setup
1. Create sensitivity labels
2. Create DLP policy with rules
3. Create alert policy for violations
4. Apply retention policy

### Secure External Access
1. Create Conditional Access policy requiring MFA
2. Create SharePoint sharing policy
3. Create Teams meeting policy
4. Create alert policy for monitoring

### Device Compliance Setup
1. Create ActiveSync mailbox policy
2. Create Conditional Access policy requiring compliant device
3. Create SharePoint access policy for unmanaged devices
4. Create alert policy for non-compliant access attempts

### Email Security Setup
1. Create Safe Attachments policy
2. Create Safe Links policy
3. Create Anti-Phishing policy
4. Create Anti-Spam policy
5. Create alert policy for security events

## Policy IDs and Common Values

### Common Conditional Access Role IDs
- Global Administrator: `62e90394-69f5-4237-9190-012177145e10`
- Security Administrator: `194ae4cb-b126-40b2-bd5b-6091b380977d`
- Exchange Administrator: `29232cdf-9323-42fd-ade2-1d097af3e4de`
- SharePoint Administrator: `f28a1f50-f6e7-4571-818b-6a12f2af6b6c`
- Teams Administrator: `69091246-20e8-4a56-aa4d-066075b2a7a8`

### Common Application IDs
- Office 365: `00000003-0000-0ff1-ce00-000000000000`
- Microsoft Teams: `cc15fd57-2c6c-4117-a88c-83b1d56b4bbe`
- Office 365 SharePoint Online: `00000003-0000-0ff1-ce00-000000000000`

### Common Location Names
- All trusted locations: `AllTrusted`
- All locations: `All`
- MFA Trusted IPs: `MfaTrustedIps`

## Error Handling

All policy management tools return consistent error messages:

- **Invalid parameters**: Missing required fields or invalid values
- **Permission denied**: Insufficient Graph API permissions
- **Resource not found**: Policy ID doesn't exist
- **Conflict**: Policy name already exists or conflicts with existing policy

Check the error message for specific details and required permissions.

## Best Practices

1. **Start with Test Mode** - Use `state: "enabledForReportingButNotEnforced"` for CA policies
2. **Use Descriptions** - Always include clear descriptions for policies
3. **Test Before Production** - Create and test policies in a dev tenant first
4. **Document Assignments** - Keep track of which policies are assigned to which users/groups
5. **Regular Reviews** - Periodically review and update policies
6. **Monitor Alerts** - Set up alert policies to monitor policy effectiveness
7. **Version Control** - Keep backups of policy configurations
8. **Gradual Rollout** - Roll out policies gradually to avoid disruption
9. **User Communication** - Inform users about new policies before enforcement
10. **Compliance Tracking** - Use reports to track policy compliance

## Troubleshooting

### Common Issues

**Issue**: "Permission denied" error
- **Solution**: Verify the app registration has required Graph API permissions

**Issue**: Policy not applying
- **Solution**: Check policy state (enabled vs disabled) and assignments

**Issue**: Conditional Access policy blocking access
- **Solution**: Review policy conditions and exclusions, use break glass accounts

**Issue**: DLP policy blocking legitimate activity
- **Solution**: Refine policy conditions, use test mode first

**Issue**: Teams policy not taking effect
- **Solution**: Wait 24-48 hours for policy propagation, verify assignment

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Conditional Access Documentation](https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/)
- [Microsoft Purview Documentation](https://docs.microsoft.com/en-us/microsoft-365/compliance/)
- [Defender for Office 365 Documentation](https://docs.microsoft.com/en-us/microsoft-365/security/office-365-security/)
- [Teams Policy Documentation](https://docs.microsoft.com/en-us/microsoftteams/policy-assignment-overview)

---

For detailed API documentation and advanced scenarios, see the complete implementation guide: `POLICY_MANAGEMENT_EXPANSION_COMPLETE.md`