// Intune macOS Device Management Types
export interface IntuneMacOSDeviceArgs {
  action: 'list' | 'get' | 'enroll' | 'retire' | 'wipe' | 'restart' | 'sync' | 'remote_lock' | 'collect_logs';
  deviceId?: string;
  filter?: string;
  enrollmentType?: 'UserEnrollment' | 'DeviceEnrollment' | 'AutomaticDeviceEnrollment';
  assignmentTarget?: {
    groupIds?: string[];
    userIds?: string[];
    deviceIds?: string[];
  };
}

export interface MacOSDevice {
  id: string;
  deviceName: string;
  managedDeviceName: string;
  userId: string;
  userDisplayName: string;
  userPrincipalName: string;
  deviceRegistrationState: 'notRegistered' | 'registered' | 'revoked' | 'keyConflict' | 'approvalPending' | 'certificateReset';
  managementState: 'managed' | 'retirePending' | 'retireIssued' | 'wipePending' | 'wipeIssued' | 'unhealthy' | 'deletePending' | 'retireFailed' | 'wipeFailed' | 'attention';
  enrolledDateTime: string;
  lastSyncDateTime: string;
  operatingSystem: string;
  osVersion: string;
  serialNumber: string;
  manufacturer: string;
  model: string;
  complianceState: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown' | 'inGracePeriod';
  jailBroken: string;
  managementAgent: 'eas' | 'mdm' | 'easMdm' | 'intuneClient' | 'easIntuneClient' | 'configurationManagerClient' | 'configurationManagerClientMdm' | 'configurationManagerClientMdmEas' | 'unknown' | 'jamf' | 'googleCloudDevicePolicyController';
  enrollmentProfileName?: string;
  bootstrapTokenEscrowed: boolean;
  activationLockBypassCode?: string;
}

// Intune macOS Policy Management Types
export interface IntuneMacOSPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'assign' | 'deploy';
  policyId?: string;
  policyType: 'Configuration' | 'Compliance' | 'Security' | 'Update' | 'AppProtection';
  name?: string;
  description?: string;
  settings?: MacOSPolicySettings;
  assignments?: PolicyAssignment[];
  deploymentSettings?: {
    installBehavior?: 'doNotInstall' | 'installAsManaged' | 'installAsUnmanaged';
    uninstallOnDeviceRemoval?: boolean;
    installAsManaged?: boolean;
  };
}

export interface MacOSPolicySettings {
  // Configuration Profile Settings
  restrictions?: MacOSRestrictions;
  security?: MacOSSecuritySettings;
  systemConfiguration?: MacOSSystemConfiguration;
  networking?: MacOSNetworkingSettings;
  applications?: MacOSApplicationSettings;
  
  // Compliance Settings
  compliance?: MacOSComplianceSettings;
  
  // Custom Settings
  customConfiguration?: {
    payloadFileName: string;
    payload: string; // Base64 encoded .mobileconfig file
  };
}

export interface MacOSRestrictions {
  allowAppInstallation?: boolean;
  allowAppRemoval?: boolean;
  allowSystemUIAppRemoval?: boolean;
  allowUniversalAppInstallation?: boolean;
  allowGuestUser?: boolean;
  allowPasswordAutoFill?: boolean;
  allowPasswordProximityRequests?: boolean;
  allowPasswordSharing?: boolean;
  allowSafariAutoFill?: boolean;
  allowScreenCapture?: boolean;
  allowRemoteScreenObservation?: boolean;
  allowAirDrop?: boolean;
  allowAirPlayIncomingRequests?: boolean;
  allowBluetoothModification?: boolean;
  allowDefinitionLookup?: boolean;
  allowFingerprintForUnlock?: boolean;
  allowSpotlightInternetResults?: boolean;
  allowTouchIdForUnlock?: boolean;
  allowWallpaperModification?: boolean;
  requirePasswordAfterScreensaver?: boolean;
  requireAdminPasswordToDeleteSystemApps?: boolean;
}

export interface MacOSSecuritySettings {
  filevault?: {
    enabled: boolean;
    recoveryKeyRotation?: boolean;
    hideRecoveryKey?: boolean;
    personalRecoveryKeyHelpMessage?: string;
    allowFDEDisableUserInitiated?: boolean;
  };
  firewall?: {
    enabled: boolean;
    blockAllIncoming?: boolean;
    enableStealthMode?: boolean;
  };
  gatekeeper?: {
    enableGatekeeper?: boolean;
    allowIdentifiedDevelopers?: boolean;
    enableGatekeeperAssessment?: boolean;
  };
  systemIntegrityProtection?: {
    enabled: boolean;
  };
  secureBootLevel?: 'off' | 'medium' | 'full';
  allowedKernelExtensions?: {
    teamIdentifier: string;
    bundleIdentifiers: string[];
  }[];
  allowedSystemExtensions?: {
    teamIdentifier: string;
    allowedTypes: ('driverExtension' | 'networkExtension' | 'endpointSecurityExtension')[];
    bundleIdentifiers: string[];
  }[];
}

export interface MacOSSystemConfiguration {
  computerName?: string;
  hostName?: string;
  localHostName?: string;
  systemPreferences?: {
    dock?: DockSettings;
    energySaver?: EnergySaverSettings;
    loginWindow?: LoginWindowSettings;
    timeZone?: string;
    networkTimeServer?: string;
  };
}

export interface DockSettings {
  dockSize?: number;
  magnification?: boolean;
  largeSize?: number;
  orientation?: 'bottom' | 'left' | 'right';
  minimizeToAppIcon?: boolean;
  showRecentAppsInDock?: boolean;
  launchpadSize?: number;
  staticItems?: DockItem[];
}

export interface DockItem {
  type: 'application' | 'directory' | 'url';
  path?: string;
  url?: string;
  label: string;
}

export interface EnergySaverSettings {
  desktopACPowerSettings?: PowerSettings;
  desktopBatteryPowerSettings?: PowerSettings;
  portableACPowerSettings?: PowerSettings;
  portableBatteryPowerSettings?: PowerSettings;
}

export interface PowerSettings {
  systemSleepTimer?: number;
  displaySleepTimer?: number;
  diskSleepTimer?: number;
  automaticRestartOnPowerLoss?: boolean;
  wakeOnLAN?: boolean;
  wakeOnRing?: boolean;
}

export interface LoginWindowSettings {
  loginWindowText?: string;
  shutDownDisabled?: boolean;
  restartDisabled?: boolean;
  sleepDisabled?: boolean;
  disableConsoleAccess?: boolean;
  loginWindowLaunchApplication?: string;
  hideLocalUsers?: boolean;
  includeNetworkUser?: boolean;
  hideAdminUsers?: boolean;
  hideMobileAccounts?: boolean;
  showFullName?: boolean;
  hideOtherUsers?: boolean;
  showPasswordHints?: boolean;
  allowGuestUser?: boolean;
  allowAutomaticLogin?: boolean;
}

export interface MacOSNetworkingSettings {
  globalHttpProxy?: ProxySettings;
  globalHttpsProxy?: ProxySettings;
  proxyCaptiveLoginAllowed?: boolean;
  vpnConfiguration?: VPNConfiguration[];
  wifiProfiles?: WiFiProfile[];
  certificateProfiles?: CertificateProfile[];
}

export interface ProxySettings {
  enabled: boolean;
  server?: string;
  port?: number;
  username?: string;
  password?: string;
  bypassDomainsAndAddresses?: string[];
}

export interface VPNConfiguration {
  connectionName: string;
  connectionType: 'IKEv2' | 'IPSec' | 'L2TP' | 'PPTP' | 'Cisco';
  server: string;
  account?: string;
  authenticationMethod: 'usernameAndPassword' | 'certificate' | 'sharedSecret';
  sharedSecret?: string;
  certificateIdentifier?: string;
  onDemandEnabled?: boolean;
  onDemandRules?: VPNOnDemandRule[];
}

export interface VPNOnDemandRule {
  action: 'connect' | 'disconnect' | 'evaluateConnection' | 'ignore';
  dnsSearchDomains?: string[];
  dnsServers?: string[];
  domainAction?: 'connectIfNeeded' | 'neverConnect';
  domains?: string[];
  requiredDNSServers?: string[];
  requiredURLStringProbe?: string;
}

export interface WiFiProfile {
  ssid: string;
  hiddenNetwork?: boolean;
  autoJoin?: boolean;
  security?: 'none' | 'wep' | 'wpa' | 'wpa2' | 'wpa3' | 'any';
  password?: string;
  eapSettings?: EAPSettings;
  proxySettings?: ProxySettings;
}

export interface EAPSettings {
  acceptedEAPTypes: number[];
  username?: string;
  outerIdentity?: string;
  password?: string;
  certificateIdentifier?: string;
  trustedCertificates?: string[];
  tlsAllowTrustExceptions?: boolean;
}

export interface CertificateProfile {
  certificateName: string;
  certificateTemplateName?: string;
  certificateAuthority?: string;
  renewalThresholdPercentage?: number;
  keySize?: 1024 | 2048 | 4096;
  keyUsage?: number;
  subjectAlternativeNameType?: 'none' | 'emailAddress' | 'userPrincipalName' | 'customAzureADAttribute' | 'domainNameService';
  subjectNameFormat?: 'commonName' | 'commonNameIncludingEmail' | 'commonNameAsEmail' | 'custom' | 'commonNameAsIMEI' | 'commonNameAsSerialNumber';
}

export interface MacOSApplicationSettings {
  managedApps?: ManagedApp[];
  allowedApps?: string[];
  blockedApps?: string[];
  appInstallationPolicy?: 'notConfigured' | 'allowList' | 'blockList';
  appAutoUpdatePolicy?: 'notConfigured' | 'enabled' | 'disabled';
}

export interface ManagedApp {
  bundleIdentifier: string;
  appName: string;
  publisher: string;
  minimumSupportedOperatingSystem: string;
  installBehavior: 'doNotInstall' | 'installAsManaged' | 'installAsUnmanaged';
  uninstallOnDeviceRemoval: boolean;
  appConfiguration?: Record<string, any>;
}

export interface MacOSComplianceSettings {
  passwordRequired?: boolean;
  passwordMinimumLength?: number;
  passwordMinutesOfInactivityBeforeLock?: number;
  passwordMinutesOfInactivityBeforeScreenTimeout?: number;
  passwordPreviousPasswordBlockCount?: number;
  passwordRequiredType?: 'deviceDefault' | 'alphanumeric' | 'numeric';
  passwordRequireToUnlockFromIdle?: boolean;
  deviceThreatProtectionEnabled?: boolean;
  deviceThreatProtectionRequiredSecurityLevel?: 'unavailable' | 'secured' | 'low' | 'medium' | 'high' | 'notSet';
  storageRequireEncryption?: boolean;
  osMinimumVersion?: string;
  osMaximumVersion?: string;
  systemIntegrityProtectionEnabled?: boolean;
  firewallEnabled?: boolean;
  firewallBlockAllIncoming?: boolean;
  firewallEnableStealthMode?: boolean;
  gatekeeperAllowedAppSource?: 'notConfigured' | 'macAppStore' | 'macAppStoreAndIdentifiedDevelopers' | 'anywhere';
  secureBootEnabled?: boolean;
  codeIntegrityEnabled?: boolean;
  advancedThreatProtectionRequiredSecurityLevel?: 'unavailable' | 'secured' | 'low' | 'medium' | 'high' | 'notSet';
}

// Policy Assignment Types
export interface PolicyAssignment {
  target: AssignmentTarget;
  intent?: 'apply' | 'remove';
  settings?: AssignmentSettings;
}

export interface AssignmentTarget {
  deviceAndAppManagementAssignmentFilterId?: string;
  deviceAndAppManagementAssignmentFilterType?: 'none' | 'include' | 'exclude';
  groupId?: string;
  collectionId?: string;
}

export interface AssignmentSettings {
  installIntent?: 'available' | 'required' | 'uninstall' | 'availableWithoutEnrollment';
  notificationSettings?: {
    showInCompanyPortal?: boolean;
    showInNotificationCenter?: boolean;
    alertType?: 'showAll' | 'showRebootsOnly' | 'hideAll';
  };
  restartSettings?: {
    restartNotificationSnoozeDurationInMinutes?: number;
    restartCountdownDisplayDurationInMinutes?: number;
  };
  deadlineSettings?: {
    useLocalTime?: boolean;
    deadlineDateTime?: string;
    gracePeriodInMinutes?: number;
  };
}

// Intune macOS App Management Types
export interface IntuneMacOSAppArgs {
  action: 'list' | 'get' | 'deploy' | 'update' | 'remove' | 'sync_status';
  appId?: string;
  appType?: 'webApp' | 'officeSuiteApp' | 'microsoftEdgeApp' | 'microsoftDefenderApp' | 'managedIOSApp' | 'managedAndroidApp' | 'managedMobileLobApp' | 'macOSLobApp' | 'macOSMicrosoftEdgeApp' | 'macOSMicrosoftDefenderApp' | 'macOSOfficeSuiteApp' | 'macOSWebClip' | 'managedApp';
  assignment?: {
    groupIds: string[];
    installIntent: 'available' | 'required' | 'uninstall' | 'availableWithoutEnrollment';
    deliveryOptimizationPriority?: 'notConfigured' | 'foreground';
  };
  appInfo?: {
    displayName: string;
    description?: string;
    publisher: string;
    bundleId?: string;
    buildNumber?: string;
    versionNumber?: string;
    packageFilePath?: string;
    minimumSupportedOperatingSystem?: string;
    ignoreVersionDetection?: boolean;
    installAsManaged?: boolean;
  };
}

// Intune Compliance Monitoring Types
export interface IntuneMacOSComplianceArgs {
  action: 'get_status' | 'get_details' | 'update_policy' | 'force_evaluation';
  deviceId?: string;
  policyId?: string;
  complianceData?: {
    passwordCompliant?: boolean;
    encryptionCompliant?: boolean;
    osVersionCompliant?: boolean;
    threatProtectionCompliant?: boolean;
    systemIntegrityCompliant?: boolean;
    firewallCompliant?: boolean;
    gatekeeperCompliant?: boolean;
    jailbrokenCompliant?: boolean;
  };
}

export interface MacOSComplianceStatus {
  deviceId: string;
  deviceName: string;
  complianceState: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown' | 'inGracePeriod';
  lastReportedDateTime: string;
  userPrincipalName: string;
  complianceGracePeriodExpirationDateTime?: string;
  deviceType: string;
  osVersion: string;
  compliancePolicyDetails: CompliancePolicyDetail[];
}

export interface CompliancePolicyDetail {
  policyId: string;
  policyName: string;
  complianceState: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown' | 'inGracePeriod';
  lastReportedDateTime: string;
  settingStates: ComplianceSettingState[];
}

export interface ComplianceSettingState {
  setting: string;
  settingName: string;
  instanceDisplayName?: string;
  state: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown' | 'inGracePeriod';
  errorCode?: string;
  errorDescription?: string;
  userId?: string;
  userEmail?: string;
  currentValue?: any;
  sources: ComplianceSettingSource[];
}

export interface ComplianceSettingSource {
  id: string;
  displayName: string;
  sourceType: 'deviceConfiguration' | 'deviceCompliance' | 'deviceIntent' | 'deviceInventory' | 'deviceShellScript' | 'unknown';
}

// Windows Intune Device Management Types
export interface IntuneWindowsDeviceArgs {
  action: 'list' | 'get' | 'enroll' | 'retire' | 'wipe' | 'restart' | 'sync' | 'remote_lock' | 'collect_logs' | 'bitlocker_recovery' | 'autopilot_reset';
  deviceId?: string;
  filter?: string;
  enrollmentType?: 'AzureADJoin' | 'HybridAzureADJoin' | 'AutoPilot' | 'BulkEnrollment';
  assignmentTarget?: {
    groupIds?: string[];
    userIds?: string[];
    deviceIds?: string[];
  };
  bitlockerSettings?: {
    requireBitlockerEncryption?: boolean;
    allowBitlockerRecoveryKeyBackup?: boolean;
  };
}

export interface WindowsDevice {
  id: string;
  deviceName: string;
  managedDeviceName: string;
  userId: string;
  userDisplayName: string;
  userPrincipalName: string;
  deviceRegistrationState: 'notRegistered' | 'registered' | 'revoked' | 'keyConflict' | 'approvalPending' | 'certificateReset';
  managementState: 'managed' | 'retirePending' | 'retireIssued' | 'wipePending' | 'wipeIssued' | 'unhealthy' | 'deletePending' | 'retireFailed' | 'wipeFailed' | 'attention';
  enrolledDateTime: string;
  lastSyncDateTime: string;
  operatingSystem: string;
  osVersion: string;
  windowsOsEdition: string;
  serialNumber: string;
  manufacturer: string;
  model: string;
  complianceState: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown' | 'inGracePeriod';
  jailBroken: string;
  managementAgent: 'mdm' | 'configurationManagerClient' | 'configurationManagerClientMdm' | 'intuneClient' | 'unknown';
  enrollmentProfileName?: string;
  azureADDeviceId?: string;
  deviceEnrollmentType: 'userEnrollment' | 'deviceEnrollment' | 'appleBulkWithUser' | 'appleBulkWithoutUser' | 'windowsAzureADJoin' | 'windowsBulkUserless' | 'windowsAutoEnrollment' | 'windowsBulkAzureDomainJoin' | 'windowsCoManagement';
  activationLockBypassCode?: string;
  emailAddress?: string;
  azureActiveDirectoryDeviceId?: string;
  azureADRegistered?: boolean;
  deviceCategoryDisplayName?: string;
  isEncrypted?: boolean;
  userType?: 'member' | 'guest';
}

// Windows Policy Management Types
export interface IntuneWindowsPolicyArgs {
  action: 'list' | 'get' | 'create' | 'update' | 'delete' | 'assign' | 'deploy';
  policyId?: string;
  policyType: 'Configuration' | 'Compliance' | 'Security' | 'Update' | 'AppProtection' | 'EndpointSecurity';
  name?: string;
  description?: string;
  settings?: WindowsPolicySettings;
  assignments?: PolicyAssignment[];
  deploymentSettings?: {
    installBehavior?: 'doNotInstall' | 'installAsManaged' | 'installAsUnmanaged';
    uninstallOnDeviceRemoval?: boolean;
    installAsManaged?: boolean;
    rebootBehavior?: 'allow' | 'suppress' | 'force';
  };
}

export interface WindowsPolicySettings {
  // Device Configuration Settings
  restrictions?: WindowsRestrictions;
  security?: WindowsSecuritySettings;
  systemConfiguration?: WindowsSystemConfiguration;
  networking?: WindowsNetworkingSettings;
  applications?: WindowsApplicationSettings;
  
  // Compliance Settings
  compliance?: WindowsComplianceSettings;
  
  // Windows Update Settings
  windowsUpdate?: WindowsUpdateSettings;
  
  // BitLocker Settings
  bitLocker?: BitLockerSettings;
  
  // Custom Settings
  customConfiguration?: {
    templateId?: string;
    settingDefinitions: CustomSetting[];
  };
}

export interface WindowsRestrictions {
  allowCortana?: boolean;
  allowSearchToUseLocation?: boolean;
  allowTaskManager?: boolean;
  allowPowerShell?: boolean;
  allowManualRootCertificateInstallation?: boolean;
  allowAddProvisioningPackage?: boolean;
  allowRemoveProvisioningPackage?: boolean;
  allowCamera?: boolean;
  allowScreenCapture?: boolean;
  allowVoiceRecording?: boolean;
  allowSyncMySettings?: boolean;
  allowWindowsSpotlight?: boolean;
  allowActionCenterNotifications?: boolean;
  allowToasts?: boolean;
  allowManualMDMUnenrollment?: boolean;
  allowAutomaticReDeployment?: boolean;
  allowSystemReset?: boolean;
  allowDeviceReset?: boolean;
  allowPasswordReveal?: boolean;
  allowPinLogon?: boolean;
  allowFingerprintUnlock?: boolean;
  microsoftAccountSignInAssistantSettings?: 'notConfigured' | 'disabled';
  antiTheftModeBlocked?: boolean;
  uninstallBuiltInApps?: boolean;
  configureTimeZone?: string;
}

export interface WindowsSecuritySettings {
  bitLockerSystemDrivePolicy?: {
    startupAuthenticationRequired?: boolean;
    startupAuthenticationBlockWithoutTpmChip?: boolean;
    startupAuthenticationTpmUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    startupAuthenticationTpmPinUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    startupAuthenticationTpmKeyUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    startupAuthenticationTpmPinAndKeyUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    minimumPinLength?: number;
    recoveryOptions?: BitLockerRecoveryOptions;
    prebootRecoveryEnableMessageAndUrl?: boolean;
    prebootRecoveryMessage?: string;
    prebootRecoveryUrl?: string;
  };
  firewallProfileDomain?: WindowsFirewallProfile;
  firewallProfilePublic?: WindowsFirewallProfile;
  firewallProfilePrivate?: WindowsFirewallProfile;
  defenderSettings?: {
    enableRealTimeProtection?: boolean;
    enableCloudProtection?: boolean;
    enableAutomaticSampleSubmission?: boolean;
    signatureUpdateIntervalInHours?: number;
  };
}

export interface BitLockerRecoveryOptions {
  blockDataRecoveryAgent?: boolean;
  recoveryPasswordUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
  recoveryKeyUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
  hideRecoveryOptions?: boolean;
  enableRecoveryInformationSaveToStore?: boolean;
  recoveryInformationToStore?: 'passwordAndKey' | 'passwordOnly' | 'keyOnly';
  enableBitLockerAfterRecoveryInformationToStore?: boolean;
}

export interface WindowsFirewallProfile {
  policyRulesFromGroupPolicyMerged?: boolean;
  policyRulesFromGroupPolicyNotMerged?: boolean;
  firewallEnabled?: 'notConfigured' | 'blocked' | 'allowed';
  stealthModeRequired?: boolean;
  stealthModeBlocked?: boolean;
  incomingTrafficRequired?: boolean;
  incomingTrafficBlocked?: boolean;
  unicastResponsesToMulticastBroadcastsRequired?: boolean;
  unicastResponsesToMulticastBroadcastsBlocked?: boolean;
  inboundNotificationsRequired?: boolean;
  inboundNotificationsBlocked?: boolean;
  authorizedApplicationRulesFromGroupPolicyMerged?: boolean;
  authorizedApplicationRulesFromGroupPolicyNotMerged?: boolean;
  globalPortRulesFromGroupPolicyMerged?: boolean;
  globalPortRulesFromGroupPolicyNotMerged?: boolean;
  connectionSecurityRulesFromGroupPolicyMerged?: boolean;
  connectionSecurityRulesFromGroupPolicyNotMerged?: boolean;
  outboundConnectionsRequired?: boolean;
  outboundConnectionsBlocked?: boolean;
  inboundConnectionsRequired?: boolean;
  inboundConnectionsBlocked?: boolean;
  securedPacketExemptionAllowed?: boolean;
  securedPacketExemptionBlocked?: boolean;
}

export interface WindowsSystemConfiguration {
  powerSettings?: {
    hibernateOnBattery?: 'notConfigured' | 'enabled' | 'disabled';
    hibernateOnPower?: 'notConfigured' | 'enabled' | 'disabled';
    standbyTimeoutOnBattery?: number;
    standbyTimeoutOnPower?: number;
    lidCloseActionOnBattery?: 'notConfigured' | 'noAction' | 'sleep' | 'hibernate' | 'shutdown';
    lidCloseActionOnPower?: 'notConfigured' | 'noAction' | 'sleep' | 'hibernate' | 'shutdown';
    powerButtonActionOnBattery?: 'notConfigured' | 'noAction' | 'sleep' | 'hibernate' | 'shutdown';
    powerButtonActionOnPower?: 'notConfigured' | 'noAction' | 'sleep' | 'hibernate' | 'shutdown';
    sleepButtonActionOnBattery?: 'notConfigured' | 'noAction' | 'sleep' | 'hibernate' | 'shutdown';
    sleepButtonActionOnPower?: 'notConfigured' | 'noAction' | 'sleep' | 'hibernate' | 'shutdown';
  };
  startMenuSettings?: {
    settingsBlockSettingsApp?: boolean;
    settingsBlockSystemPage?: boolean;
    settingsBlockDevicesPage?: boolean;
    settingsBlockNetworkInternetPage?: boolean;
    settingsBlockPersonalizationPage?: boolean;
    settingsBlockAccountsPage?: boolean;
    settingsBlockTimeLanguagePage?: boolean;
    settingsBlockEaseOfAccessPage?: boolean;
    settingsBlockPrivacyPage?: boolean;
    settingsBlockUpdateSecurityPage?: boolean;
    settingsBlockAppsPage?: boolean;
    settingsBlockGamingPage?: boolean;
    windowsSpotlightBlockConsumerSpecificFeatures?: boolean;
    windowsSpotlightBlocked?: boolean;
    windowsSpotlightBlockOnActionCenter?: boolean;
    windowsSpotlightBlockTailoredExperiences?: boolean;
    windowsSpotlightBlockThirdPartyNotifications?: boolean;
    windowsSpotlightBlockWelcomeExperience?: boolean;
    windowsSpotlightBlockWindowsTips?: boolean;
    windowsSpotlightConfigureOnLockScreen?: 'notConfigured' | 'disabled' | 'enabled';
  };
}

export interface WindowsNetworkingSettings {
  networkIsolationPolicy?: {
    enterpriseNetworkDomainNames?: string[];
    enterpriseCloudResources?: CloudResource[];
    enterpriseInternalProxyServers?: string[];
    enterpriseIPRanges?: IPRange[];
    enterpriseIsAuthorized?: boolean;
    enterpriseProxyServers?: string[];
    enterpriseProxyServersAreAuthoritative?: boolean;
    neutralDomainResources?: string[];
  };
  vpnConfiguration?: {
    connectionName?: string;
    servers?: VpnServer[];
    customXml?: string;
  };
}

export interface CloudResource {
  displayName?: string;
  resources?: string[];
}

export interface IPRange {
  displayName?: string;
  lowerAddress?: string;
  upperAddress?: string;
}

export interface VpnServer {
  description?: string;
  address?: string;
  isDefaultServer?: boolean;
}

export interface WindowsApplicationSettings {
  appsBlockWindowsStoreOriginatedApps?: boolean;
  appsAllowTrustedAppsSideloading?: 'notConfigured' | 'blocked' | 'allowed';
  gameDvrBlocked?: boolean;
  experienceDoNotSyncBrowserSettings?: 'notConfigured' | 'blockedSynchronization' | 'allowedSynchronization';
  antiTheftModeBlocked?: boolean;
  smartScreenEnabled?: boolean;
  smartScreenBlockPromptOverride?: boolean;
  smartScreenBlockPromptOverrideForFiles?: boolean;
  webRtcBlockLocalhostIpAddress?: boolean;
  internetSharingBlocked?: boolean;
  settingsBlockAddProvisioningPackage?: boolean;
  settingsBlockRemoveProvisioningPackage?: boolean;
  settingsBlockChangeSystemTime?: boolean;
  settingsBlockEditDeviceName?: boolean;
  settingsBlockChangeRegion?: boolean;
  settingsBlockChangeLanguage?: boolean;
  settingsBlockChangePowerSleep?: boolean;
  locationServicesBlocked?: boolean;
  microsoftAccountBlocked?: boolean;
  microsoftAccountBlockSettingsSync?: boolean;
  nfcBlocked?: boolean;
  resetProtectionModeBlocked?: boolean;
  screenCaptureBlocked?: boolean;
  storageBlockRemovableStorage?: boolean;
  storageRequireMobileDeviceEncryption?: boolean;
  usbBlocked?: boolean;
  voiceRecordingBlocked?: boolean;
  wiFiBlockAutomaticConnectHotspots?: boolean;
  wiFiBlocked?: boolean;
  wiFiBlockManualConfiguration?: boolean;
  wiFiScanInterval?: number;
  bluetoothAllowedServices?: string[];
  bluetoothBlockAdvertising?: boolean;
  bluetoothBlockPromptedProximalConnections?: boolean;
  bluetoothBlockDiscoverableMode?: boolean;
  bluetoothBlockPrePairing?: boolean;
  edgeBlocked?: boolean;
  edgeCookiePolicy?: 'userDefined' | 'allow' | 'blockThirdParty' | 'blockAll';
  edgeBlockPasswordManager?: boolean;
  edgeBlockAddressBarDropdown?: boolean;
  edgeBlockCompatibilityList?: boolean;
  edgeClearBrowsingDataOnExit?: boolean;
  edgeAllowStartPagesModification?: boolean;
  edgeDisableFirstRunPage?: boolean;
  edgeBlockLiveTileDataCollection?: boolean;
  edgeSyncFavoritesWithInternetExplorer?: boolean;
}

export interface WindowsComplianceSettings {
  passwordRequired?: boolean;
  passwordBlockSimple?: boolean;
  passwordExpirationDays?: number;
  passwordMinimumLength?: number;
  passwordMinutesOfInactivityBeforeLock?: number;
  passwordMinimumCharacterSetCount?: number;
  passwordPreviousPasswordBlockCount?: number;
  passwordRequiredType?: 'deviceDefault' | 'alphanumeric' | 'numeric';
  requireHealthyDeviceReport?: boolean;
  osMinimumVersion?: string;
  osMaximumVersion?: string;
  mobileOsMinimumVersion?: string;
  mobileOsMaximumVersion?: string;
  earlyLaunchAntiMalwareDriverEnabled?: boolean;
  bitLockerEnabled?: boolean;
  secureBootEnabled?: boolean;
  codeIntegrityEnabled?: boolean;
  storageRequireEncryption?: boolean;
  activeFirewallRequired?: boolean;
  defenderEnabled?: boolean;
  defenderVersion?: string;
  signatureOutOfDate?: boolean;
  rtpEnabled?: boolean;
  antivirusRequired?: boolean;
  antiSpywareRequired?: boolean;
}

export interface WindowsUpdateSettings {
  deliveryOptimizationMode?: 'userDefined' | 'httpOnly' | 'httpWithPeeringNat' | 'httpWithPeeringPrivateGroup' | 'httpWithInternetPeering' | 'simpleDownload' | 'bypassMode';
  prereleaseFeatures?: 'userDefined' | 'settingsOnly' | 'settingsAndExperimentations' | 'notAllowed';
  automaticUpdateMode?: 'userDefined' | 'notifyDownload' | 'autoInstallAtMaintenanceTime' | 'autoInstallAndRebootAtMaintenanceTime' | 'autoInstallAndRebootAtScheduledTime' | 'autoInstallAndRebootWithoutEndUserControl' | 'windowsDefault';
  microsoftUpdateServiceAllowed?: boolean;
  driversExcluded?: boolean;
  installationSchedule?: WindowsUpdateInstallationSchedule;
  qualityUpdatesDeferralPeriodInDays?: number;
  featureUpdatesDeferralPeriodInDays?: number;
  qualityUpdatesPaused?: boolean;
  featureUpdatesPaused?: boolean;
  qualityUpdatesPauseExpiryDateTime?: string;
  featureUpdatesPauseExpiryDateTime?: string;
  businessReadyUpdatesOnly?: 'userDefined' | 'all' | 'businessReadyOnly';
}

export interface WindowsUpdateInstallationSchedule {
  activeHoursStart?: string;
  activeHoursEnd?: string;
  scheduledInstallDay?: 'userDefined' | 'everyday' | 'sunday' | 'monday' | 'tuesday' | 'wednesday' | 'thursday' | 'friday' | 'saturday';
  scheduledInstallTime?: string;
}

export interface BitLockerSettings {
  systemDrivePolicy?: {
    encryptionMethod?: 'aesCbc128' | 'aesCbc256' | 'xtsAes128' | 'xtsAes256';
    startupAuthenticationRequired?: boolean;
    startupAuthenticationBlockWithoutTpmChip?: boolean;
    startupAuthenticationTpmUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    startupAuthenticationTpmPinUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    startupAuthenticationTpmKeyUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    startupAuthenticationTpmPinAndKeyUsage?: 'blocked' | 'required' | 'allowed' | 'notConfigured';
    minimumPinLength?: number;
    recoveryOptions?: BitLockerRecoveryOptions;
    prebootRecoveryEnableMessageAndUrl?: boolean;
    prebootRecoveryMessage?: string;
    prebootRecoveryUrl?: string;
  };
  fixedDrivePolicy?: {
    encryptionMethod?: 'aesCbc128' | 'aesCbc256' | 'xtsAes128' | 'xtsAes256';
    requireEncryptionForWriteAccess?: boolean;
    recoveryOptions?: BitLockerRecoveryOptions;
  };
  removableDrivePolicy?: {
    encryptionMethod?: 'aesCbc128' | 'aesCbc256' | 'xtsAes128' | 'xtsAes256';
    requireEncryptionForWriteAccess?: boolean;
    blockCrossOrganizationWriteAccess?: boolean;
  };
}

export interface CustomSetting {
  settingDefinitionId?: string;
  value?: any;
}

// Windows App Management Types
export interface IntuneWindowsAppArgs {
  action: 'list' | 'get' | 'deploy' | 'update' | 'remove' | 'sync_status';
  appId?: string;
  appType?: 'win32LobApp' | 'microsoftStoreForBusinessApp' | 'webApp' | 'officeSuiteApp' | 'microsoftEdgeApp';
  name?: string;
  version?: string;
  assignmentGroups?: string[];
  assignment?: {
    groupIds: string[];
    installIntent: 'available' | 'required' | 'uninstall' | 'availableWithoutEnrollment';
    deliveryOptimizationPriority?: 'notConfigured' | 'foreground';
  };
  appInfo?: {
    displayName: string;
    description?: string;
    publisher: string;
    fileName?: string;
    setupFilePath?: string;
    installCommandLine?: string;
    uninstallCommandLine?: string;
    minimumSupportedOperatingSystem?: string;
    packageFilePath?: string;
  };
}

// Windows Compliance Management Types
export interface IntuneWindowsComplianceArgs {
  action: 'get_status' | 'get_details' | 'update_policy' | 'force_evaluation' | 'get_bitlocker_keys';
  deviceId?: string;
  complianceType?: 'security' | 'configuration' | 'update' | 'bitlocker';
  policies?: string[];
  complianceData?: {
    passwordCompliant?: boolean;
    encryptionCompliant?: boolean;
    osVersionCompliant?: boolean;
    threatProtectionCompliant?: boolean;
    bitlockerCompliant?: boolean;
    firewallCompliant?: boolean;
    antivirusCompliant?: boolean;
  };
}
