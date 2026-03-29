const categoryResourceMap = {
    accessreview: 'accessReviewScheduleDefinition',
    administrativeunit: 'administrativeUnit',
    agreement: 'agreement',
    application: 'application',
    auditlog: 'directoryAudit',
    bookings: 'bookingBusiness',
    calendars: 'calendar',
    calls: 'call',
    channel: 'channel',
    chat: 'chat',
    cloudpc: 'cloudPcRoot',
    contacts: 'contact',
    device: 'device',
    directory: 'organization',
    domain: 'domain',
    ediscovery: 'ediscoveryCase',
    entitlementmanagement: 'accessPackage',
    externalconnection: 'externalConnection',
    files: 'driveItem',
    group: 'group',
    identitygovernance: 'identityGovernance',
    identityprovider: 'identityProviderBase',
    lifecycleworkflows: 'lifecycleWorkflowsContainer',
    mail: 'message',
    notes: 'onenote',
    onlinemeetings: 'onlineMeeting',
    people: 'person',
    planner: 'plannerTask',
    policy: 'policyRoot',
    presence: 'presence',
    printer: 'printer',
    reports: 'reportRoot',
    rolemanagement: 'unifiedRoleDefinition',
    schedule: 'schedule',
    securityevents: 'alert',
    serviceprincipal: 'servicePrincipal',
    sites: 'site',
    synchronization: 'synchronization',
    tasks: 'todoTask',
    teams: 'team',
    threatassessment: 'threatAssessmentRequest',
    threatintelligence: 'threatIntelligence',
    user: 'user',
    virtualevent: 'virtualEvent'
};

const relationshipsMap = {
    accessreview: [
        { name: 'instances', type: 'accessReviewInstance collection', description: 'Instances created for the access review definition.' }
    ],
    administrativeunit: [
        { name: 'members', type: 'directoryObject collection', description: 'Users and groups that are members of this administrative unit.' },
        { name: 'scopedRoleMembers', type: 'scopedRoleMembership collection', description: 'Scoped-role members of this administrative unit.' }
    ],
    agreement: [
        { name: 'acceptances', type: 'agreementAcceptance collection', description: 'Acceptances recorded for this agreement.' }
    ],
    application: [
        { name: 'owners', type: 'directoryObject collection', description: 'Directory objects that are owners of the application.' },
        { name: 'federatedIdentityCredentials', type: 'federatedIdentityCredential collection', description: 'Federated identities configured for the application.' },
        { name: 'extensionProperties', type: 'extensionProperty collection', description: 'Extension properties defined on the application.' }
    ],
    bookings: [
        { name: 'appointments', type: 'bookingAppointment collection', description: 'Appointments for this business.' },
        { name: 'customers', type: 'bookingCustomerBase collection', description: 'Customers of this business.' },
        { name: 'services', type: 'bookingService collection', description: 'Services offered by this business.' },
        { name: 'staffMembers', type: 'bookingStaffMemberBase collection', description: 'Staff members assigned to this business.' }
    ],
    calendars: [
        { name: 'calendarView', type: 'event collection', description: 'Calendar view for the calendar.' },
        { name: 'events', type: 'event collection', description: 'Events in the calendar.' },
        { name: 'calendarPermissions', type: 'calendarPermission collection', description: 'Sharing permissions for the calendar.' }
    ],
    calls: [
        { name: 'participants', type: 'participant collection', description: 'Participants associated with the call.' },
        { name: 'operations', type: 'commsOperation collection', description: 'Operations associated with the call.' }
    ],
    channel: [
        { name: 'members', type: 'conversationMember collection', description: 'Membership records associated with the channel.' },
        { name: 'messages', type: 'chatMessage collection', description: 'Messages in the channel.' },
        { name: 'tabs', type: 'teamsTab collection', description: 'Tabs in the channel.' },
        { name: 'filesFolder', type: 'driveItem', description: 'Folder where the channel files are stored.' }
    ],
    chat: [
        { name: 'members', type: 'conversationMember collection', description: 'Members in the chat.' },
        { name: 'messages', type: 'chatMessage collection', description: 'Messages in the chat.' },
        { name: 'installedApps', type: 'teamsAppInstallation collection', description: 'Apps installed in the chat.' },
        { name: 'tabs', type: 'teamsTab collection', description: 'Tabs in the chat.' }
    ],
    cloudpc: [
        { name: 'cloudPCs', type: 'cloudPC collection', description: 'Cloud PCs in the tenant.' },
        { name: 'provisioningPolicies', type: 'cloudPcProvisioningPolicy collection', description: 'Provisioning policies for Cloud PCs.' }
    ],
    contacts: [
        { name: 'photo', type: 'profilePhoto', description: 'Optional contact picture.' },
        { name: 'extensions', type: 'extension collection', description: 'Open extensions defined for the contact.' }
    ],
    device: [
        { name: 'registeredOwners', type: 'directoryObject collection', description: 'Owners registered for the device.' },
        { name: 'registeredUsers', type: 'directoryObject collection', description: 'Users registered for the device.' },
        { name: 'memberOf', type: 'directoryObject collection', description: 'Groups and administrative units that include this device.' }
    ],
    directory: [
        { name: 'administrativeUnits', type: 'administrativeUnit collection', description: 'Administrative units in the tenant.' },
        { name: 'deletedItems', type: 'directoryObject collection', description: 'Deleted objects in the tenant.' }
    ],
    domain: [
        { name: 'federationConfiguration', type: 'internalDomainFederation collection', description: 'Federation settings for the domain.' },
        { name: 'serviceConfigurationRecords', type: 'domainDnsRecord collection', description: 'Service configuration DNS records.' }
    ],
    ediscovery: [
        { name: 'custodians', type: 'ediscoveryCustodian collection', description: 'Custodians in the eDiscovery case.' },
        { name: 'searches', type: 'ediscoverySearch collection', description: 'Searches configured in the eDiscovery case.' }
    ],
    entitlementmanagement: [
        { name: 'accessPackages', type: 'accessPackage collection', description: 'Access packages in the entitlement management catalog.' },
        { name: 'assignmentPolicies', type: 'accessPackageAssignmentPolicy collection', description: 'Assignment policies for access packages.' }
    ],
    externalconnection: [
        { name: 'groups', type: 'externalGroup collection', description: 'External groups in the connection.' },
        { name: 'items', type: 'externalItem collection', description: 'External items in the connection.' },
        { name: 'schema', type: 'schema', description: 'Schema definition for the external connection.' }
    ],
    files: [
        { name: 'children', type: 'driveItem collection', description: 'Immediate children under the drive item.' },
        { name: 'permissions', type: 'permission collection', description: 'Permissions assigned to the drive item.' },
        { name: 'versions', type: 'driveItemVersion collection', description: 'Historical versions of the drive item.' },
        { name: 'thumbnails', type: 'thumbnailSet collection', description: 'Thumbnail sets for the drive item.' }
    ],
    group: [
        { name: 'members', type: 'directoryObject collection', description: 'Direct members of the group.' },
        { name: 'owners', type: 'directoryObject collection', description: 'Owners of the group.' },
        { name: 'drive', type: 'drive', description: 'Default drive for the group.' },
        { name: 'team', type: 'team', description: 'Team associated with the group.' },
        { name: 'sites', type: 'site collection', description: 'Sites linked to the group.' }
    ],
    identitygovernance: [
        { name: 'entitlementManagement', type: 'entitlementManagement', description: 'Entitlement management container.' },
        { name: 'accessReviews', type: 'accessReviewScheduleDefinition collection', description: 'Access review definitions in governance.' }
    ],
    lifecycleworkflows: [
        { name: 'workflows', type: 'workflow collection', description: 'Lifecycle workflow definitions.' },
        { name: 'taskDefinitions', type: 'taskDefinition collection', description: 'Task definitions used in lifecycle workflows.' }
    ],
    mail: [
        { name: 'attachments', type: 'attachment collection', description: 'Attachments on the message.' },
        { name: 'extensions', type: 'extension collection', description: 'Open extensions defined for the message.' },
        { name: 'singleValueExtendedProperties', type: 'singleValueLegacyExtendedProperty collection', description: 'Single-value extended properties for the message.' },
        { name: 'multiValueExtendedProperties', type: 'multiValueLegacyExtendedProperty collection', description: 'Multi-value extended properties for the message.' }
    ],
    notes: [
        { name: 'notebooks', type: 'notebook collection', description: 'OneNote notebooks.' },
        { name: 'pages', type: 'onenotePage collection', description: 'OneNote pages.' }
    ],
    onlinemeetings: [
        { name: 'attendanceReports', type: 'meetingAttendanceReport collection', description: 'Attendance reports for the online meeting.' },
        { name: 'recordings', type: 'callRecording collection', description: 'Recordings for the online meeting.' },
        { name: 'transcripts', type: 'callTranscript collection', description: 'Transcripts for the online meeting.' }
    ],
    people: [
        { name: 'profilePhoto', type: 'profilePhoto', description: 'Profile photo for the person.' }
    ],
    planner: [
        { name: 'details', type: 'plannerTaskDetails', description: 'Additional details about the planner task.' },
        { name: 'assignedToTaskBoardFormat', type: 'plannerAssignedToTaskBoardTaskFormat', description: 'Assigned-to board format for the task.' }
    ],
    policy: [
        { name: 'conditionalAccessPolicies', type: 'conditionalAccessPolicy collection', description: 'Conditional access policies in the tenant.' },
        { name: 'permissionGrantPolicies', type: 'permissionGrantPolicy collection', description: 'Permission grant policies in the tenant.' },
        { name: 'authenticationMethodsPolicy', type: 'authenticationMethodsPolicy', description: 'Authentication methods policy.' }
    ],
    presence: [
        { name: 'statusMessage', type: 'presenceStatusMessage', description: 'Presence status message of a user.' }
    ],
    printer: [
        { name: 'jobs', type: 'printJob collection', description: 'Print jobs queued for the printer.' },
        { name: 'shares', type: 'printerShare collection', description: 'Printer shares associated with the printer.' }
    ],
    reports: [
        { name: 'security', type: 'securityReportsRoot', description: 'Security-related reports root.' }
    ],
    rolemanagement: [
        { name: 'directory', type: 'rbacApplication', description: 'Directory role management root.' },
        { name: 'entitlementManagement', type: 'rbacApplication', description: 'Entitlement management role container.' }
    ],
    schedule: [
        { name: 'shifts', type: 'shift collection', description: 'Shifts in the schedule.' },
        { name: 'timeOffRequests', type: 'timeOffRequest collection', description: 'Time off requests in the schedule.' },
        { name: 'openShifts', type: 'openShift collection', description: 'Open shifts in the schedule.' }
    ],
    securityevents: [
        { name: 'alerts', type: 'alert collection', description: 'Security alerts in the tenant.' },
        { name: 'secureScores', type: 'secureScore collection', description: 'Secure score records in the tenant.' }
    ],
    serviceprincipal: [
        { name: 'appRoleAssignedTo', type: 'appRoleAssignment collection', description: 'App role assignments granted to users, groups or service principals.' },
        { name: 'owners', type: 'directoryObject collection', description: 'Owners of the service principal.' },
        { name: 'claimsMappingPolicies', type: 'claimsMappingPolicy collection', description: 'Claims mapping policies assigned to the service principal.' }
    ],
    sites: [
        { name: 'drive', type: 'drive', description: 'Default drive for the site.' },
        { name: 'drives', type: 'drive collection', description: 'Document libraries under the site.' },
        { name: 'lists', type: 'list collection', description: 'Lists under the site.' },
        { name: 'sites', type: 'site collection', description: 'Subsites under the site.' },
        { name: 'permissions', type: 'permission collection', description: 'Permissions associated with the site.' }
    ],
    synchronization: [
        { name: 'jobs', type: 'synchronizationJob collection', description: 'Synchronization jobs configured for the resource.' },
        { name: 'templates', type: 'synchronizationTemplate collection', description: 'Synchronization templates available for the resource.' }
    ],
    tasks: [
        { name: 'linkedResources', type: 'linkedResource collection', description: 'Resources linked to the task.' },
        { name: 'checklistItems', type: 'checklistItem collection', description: 'Checklist items linked to the task.' },
        { name: 'attachments', type: 'attachmentBase collection', description: 'Attachments linked to the task.' }
    ],
    teams: [
        { name: 'channels', type: 'channel collection', description: 'Channels associated with the team.' },
        { name: 'members', type: 'conversationMember collection', description: 'Members and owners of the team.' },
        { name: 'installedApps', type: 'teamsAppInstallation collection', description: 'Apps installed in the team.' },
        { name: 'schedule', type: 'schedule', description: 'Schedule attached to the team.' }
    ],
    threatassessment: [
        { name: 'results', type: 'threatAssessmentResult collection', description: 'Results recorded for the threat assessment.' }
    ],
    threatintelligence: [
        { name: 'articles', type: 'article collection', description: 'Threat intelligence articles.' },
        { name: 'intelProfiles', type: 'intelligenceProfile collection', description: 'Threat intelligence profiles.' }
    ],
    user: [
        { name: 'manager', type: 'directoryObject', description: 'Manager of the user.' },
        { name: 'directReports', type: 'directoryObject collection', description: 'Direct reports of the user.' },
        { name: 'memberOf', type: 'directoryObject collection', description: 'Groups, directory roles and administrative units of the user.' },
        { name: 'ownedDevices', type: 'directoryObject collection', description: 'Devices owned by the user.' },
        { name: 'drive', type: 'drive', description: 'OneDrive of the user.' },
        { name: 'calendar', type: 'calendar', description: 'Primary calendar of the user.' },
        { name: 'events', type: 'event collection', description: 'Events in the user calendar.' },
        { name: 'mailFolders', type: 'mailFolder collection', description: 'Mail folders in the user mailbox.' },
        { name: 'messages', type: 'message collection', description: 'Messages in the user mailbox.' },
        { name: 'contacts', type: 'contact collection', description: 'Contacts of the user.' },
        { name: 'photo', type: 'profilePhoto', description: 'Profile photo of the user.' },
        { name: 'planner', type: 'plannerUser', description: 'Planner services available to the user.' }
    ],
    virtualevent: [
        { name: 'events', type: 'virtualEvent collection', description: 'Virtual events in the tenant.' },
        { name: 'townhalls', type: 'virtualEventTownhall collection', description: 'Virtual event town halls.' },
        { name: 'webinars', type: 'virtualEventWebinar collection', description: 'Virtual event webinars.' }
    ]
};

function resolveResourceName(category) {
    return categoryResourceMap[category.toLowerCase()] || category;
}

function getRelationships(category) {
    return relationshipsMap[category.toLowerCase()] || [];
}

module.exports = {
    categoryResourceMap,
    getRelationships,
    relationshipsMap,
    resolveResourceName
};
