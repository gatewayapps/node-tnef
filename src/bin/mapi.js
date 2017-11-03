const utils = require('./util')

// decode any MAPI message attributes
var decodeMapi = ((data) => {
    var attrs = []
    var dataLen = data.length
    var offset = 0
    var numProperties = utils.processBytesToInteger(data, offset, 4)
    offset += 4

    for (var i = 0; i < numProperties; i++) {
        if (offset >= dataLen) {
            continue
        }

        var attrType = utils.processBytesToInteger(data, offset, 2)
        offset += 2

        var isMultiValue = (attrType & mvFlag) !== 0
        attrType &= mvFlag

        var typeSize = getTypeSize(attrType)
        if (typeSize < 0) {
            isMultiValue = true
        }

        var attrName = utils.processBytesToInteger(data, offset, 2)
        offset += 2

        var guid = 0
        if (attrName >= 0x8000 && attrName <= 0xFFFE) {
            guid = utils.processBytesToInteger(data, offset, 16)
            offset += 16
            var kind = utils.processBytesToInteger(data, offset, 4)
            offset += 4

            if (kind === 0) {
                offset += 4
            } else if (kind === 1) {
                var iidLen = utils.processBytesToInteger(data, offset, 4)
                offset += 4
                offset += iidLen
                offset += (-iidLen & 3)
            }
        }

        // handle multi-value properties
        var valueCount = 1
        if (isMultiValue) {
            valueCount = utils.processBytesToInteger(data, offset, 4)
            offset += 4
        }

        if (valueCount > 1024 && valueCount > data.length) {
            return
        }

        var attrData = []

        for (var i = 0; i < valueCount; i++) {
            var length = typeSize
            if (typeSize < 0) {
                length = utils.processBytesToInteger(data, offset, 4)
                offset += 4
            }

            // read the data in
            attrData.push(utils.processBytes(data, offset, length))

            offset += length
            offset += (-length & 3)
        }

        attrs.push({ Type: attrType, Name: attrName, Data: attrData, GUID: guid })
    }

    return attrs
})

var getTypeSize = ((attrType) => {
    switch (attrType) {
        case szmapiShort, szmapiBoolean:
            return 2
        case szmapiInt, szmapiFloat, szmapiError:
            return 4
        case szmapiDouble, szmapiApptime, szmapiCurrency, szmapiInt8byte, szmapiSystime:
            return 8
        case szmapiCLSID:
            return 16
        case szmapiString, szmapiUnicodeString, szmapiObject, szmapiBinary:
            return -1
    }
    return 0
})

const mvFlag = 0x1000 // OR with type means multiple values

const szmapiUnspecified = 0x0000 //# MAPI Unspecified
const szmapiNull = 0x0001 //# MAPI null property
const szmapiShort = 0x0002 //# MAPI short (signed 16 bits)
const szmapiInt = 0x0003 //# MAPI integer (signed 32 bits)
const szmapiFloat = 0x0004 //# MAPI float (4 bytes)
const szmapiDouble = 0x0005 //# MAPI double
const szmapiCurrency = 0x0006 //# MAPI currency (64 bits)
const szmapiApptime = 0x0007 //# MAPI application time
const szmapiError = 0x000a //# MAPI error (32 bits)
const szmapiBoolean = 0x000b //# MAPI boolean (16 bits)
const szmapiObject = 0x000d //# MAPI embedded object
const szmapiInt8byte = 0x0014 //# MAPI 8 byte signed int
const szmapiString = 0x001e //# MAPI string
const szmapiUnicodeString = 0x001f //# MAPI unicode-string (null terminated)
const szmapiPtSystime = 0x001e //# MAPI time (after 2038/01/17 22:14:07 or before 1970/01/01 00:00:00)
const szmapiSystime = 0x0040 //# MAPI time (64 bits)
const szmapiCLSID = 0x0048 //# MAPI OLE GUID
const szmapiBinary = 0x0102 //# MAPI binary
const szmapiUnknown = 0x0033

// We can use these constants to find specific types
// of MAPIAttribute by comparing it to the type of the
// attribute.

const MAPIAcknowledgementMode = 0x0001
const MAPIAlternateRecipientAllowed = 0x0002
const MAPIAuthorizingUsers = 0x0003
const MAPIAutoForwardComment = 0x0004
const MAPIAutoForwarded = 0x0005
const MAPIContentConfidentialityAlgorithmID = 0x0006
const MAPIContentCorrelator = 0x0007
const MAPIContentIdentifier = 0x0008
const MAPIContentLength = 0x0009
const MAPIContentReturnRequested = 0x000A
const MAPIConversationKey = 0x000B
const MAPIConversionEits = 0x000C
const MAPIConversionWithLossProhibited = 0x000D
const MAPIConvertedEits = 0x000E
const MAPIDeferredDeliveryTime = 0x000F
const MAPIDeliverTime = 0x0010
const MAPIDiscardReason = 0x0011
const MAPIDisclosureOfRecipients = 0x0012
const MAPIDlExpansionHistory = 0x0013
const MAPIDlExpansionProhibited = 0x0014
const MAPIExpiryTime = 0x0015
const MAPIImplicitConversionProhibited = 0x0016
const MAPIImportance = 0x0017
const MAPIIpmID = 0x0018
const MAPILatestDeliveryTime = 0x0019
const MAPIMessageClass = 0x001A
const MAPIMessageDeliveryID = 0x001B
const MAPIMessageSecurityLabel = 0x001E
const MAPIObsoletedIpms = 0x001F
const MAPIOriginallyIntendedRecipientName = 0x0020
const MAPIOriginalEits = 0x0021
const MAPIOriginatorCertificate = 0x0022
const MAPIOriginatorDeliveryReportRequested = 0x0023
const MAPIOriginatorReturnAddress = 0x0024
const MAPIParentKey = 0x0025
const MAPIPriority = 0x0026
const MAPIOriginCheck = 0x0027
const MAPIProofOfSubmissionRequested = 0x0028
const MAPIReadReceiptRequested = 0x0029
const MAPIReceiptTime = 0x002A
const MAPIRecipientReassignmentProhibited = 0x002B
const MAPIRedirectionHistory = 0x002C
const MAPIRelatedIpms = 0x002D
const MAPIOriginalSensitivity = 0x002E
const MAPILanguages = 0x002F
const MAPIReplyTime = 0x0030
const MAPIReportTag = 0x0031
const MAPIReportTime = 0x0032
const MAPIReturnedIpm = 0x0033
const MAPISecurity = 0x0034
const MAPIIncompleteCopy = 0x0035
const MAPISensitivity = 0x0036
const MAPISubject = 0x0037
const MAPISubjectIpm = 0x0038
const MAPIClientSubmitTime = 0x0039
const MAPIReportName = 0x003A
const MAPISentRepresentingSearchKey = 0x003B
const MAPIX400ContentType = 0x003C
const MAPISubjectPrefix = 0x003D
const MAPINonReceiptReason = 0x003E
const MAPIReceivedByEntryID = 0x003F
const MAPIReceivedByName = 0x0040
const MAPISentRepresentingEntryID = 0x0041
const MAPISentRepresentingName = 0x0042
const MAPIRcvdRepresentingEntryID = 0x0043
const MAPIRcvdRepresentingName = 0x0044
const MAPIReportEntryID = 0x0045
const MAPIReadReceiptEntryID = 0x0046
const MAPIMessageSubmissionID = 0x0047
const MAPIProviderSubmitTime = 0x0048
const MAPIOriginalSubject = 0x0049
const MAPIDiscVal = 0x004A
const MAPIOrigMessageClass = 0x004B
const MAPIOriginalAuthorEntryID = 0x004C
const MAPIOriginalAuthorName = 0x004D
const MAPIOriginalSubmitTime = 0x004E
const MAPIReplyRecipientEntries = 0x004F
const MAPIReplyRecipientNames = 0x0050
const MAPIReceivedBySearchKey = 0x0051
const MAPIRcvdRepresentingSearchKey = 0x0052
const MAPIReadReceiptSearchKey = 0x0053
const MAPIReportSearchKey = 0x0054
const MAPIOriginalDeliveryTime = 0x0055
const MAPIOriginalAuthorSearchKey = 0x0056
const MAPIMessageToMe = 0x0057
const MAPIMessageCcMe = 0x0058
const MAPIMessageRecipMe = 0x0059
const MAPIOriginalSenderName = 0x005A
const MAPIOriginalSenderEntryID = 0x005B
const MAPIOriginalSenderSearchKey = 0x005C
const MAPIOriginalSentRepresentingName = 0x005D
const MAPIOriginalSentRepresentingEntryID = 0x005E
const MAPIOriginalSentRepresentingSearchKey = 0x005F
const MAPIStartDate = 0x0060
const MAPIEndDate = 0x0061
const MAPIOwnerApptID = 0x0062
const MAPIResponseRequested = 0x0063
const MAPISentRepresentingAddrtype = 0x0064
const MAPISentRepresentingEmailAddress = 0x0065
const MAPIOriginalSenderAddrtype = 0x0066
const MAPIOriginalSenderEmailAddress = 0x0067
const MAPIOriginalSentRepresentingAddrtype = 0x0068
const MAPIOriginalSentRepresentingEmailAddress = 0x0069
const MAPIConversationTopic = 0x0070
const MAPIConversationIndex = 0x0071
const MAPIOriginalDisplayBcc = 0x0072
const MAPIOriginalDisplayCc = 0x0073
const MAPIOriginalDisplayTo = 0x0074
const MAPIReceivedByAddrtype = 0x0075
const MAPIReceivedByEmailAddress = 0x0076
const MAPIRcvdRepresentingAddrtype = 0x0077
const MAPIRcvdRepresentingEmailAddress = 0x0078
const MAPIOriginalAuthorAddrtype = 0x0079
const MAPIOriginalAuthorEmailAddress = 0x007A
const MAPIOriginallyIntendedRecipAddrtype = 0x007B
const MAPIOriginallyIntendedRecipEmailAddress = 0x007C
const MAPITransportMessageHeaders = 0x007D
const MAPIDelegation = 0x007E
const MAPITnefCorrelationKey = 0x007F
const MAPIBody = 0x1000
const MAPIBodyHTML = 0x1013
const MAPIReportText = 0x1001
const MAPIOriginatorAndDlExpansionHistory = 0x1002
const MAPIReportingDlName = 0x1003
const MAPIReportingMtaCertificate = 0x1004
const MAPIRtfSyncBodyCrc = 0x1006
const MAPIRtfSyncBodyCount = 0x1007
const MAPIRtfSyncBodyTag = 0x1008
const MAPIRtfCompressed = 0x1009
const MAPIRtfSyncPrefixCount = 0x1010
const MAPIRtfSyncTrailingCount = 0x1011
const MAPIOriginallyIntendedRecipEntryID = 0x1012
const MAPIContentIntegrityCheck = 0x0C00
const MAPIExplicitConversion = 0x0C01
const MAPIIpmReturnRequested = 0x0C02
const MAPIMessageToken = 0x0C03
const MAPINdrReasonCode = 0x0C04
const MAPINdrDiagCode = 0x0C05
const MAPINonReceiptNotificationRequested = 0x0C06
const MAPIDeliveryPoint = 0x0C07
const MAPIOriginatorNonDeliveryReportRequested = 0x0C08
const MAPIOriginatorRequestedAlternateRecipient = 0x0C09
const MAPIPhysicalDeliveryBureauFaxDelivery = 0x0C0A
const MAPIPhysicalDeliveryMode = 0x0C0B
const MAPIPhysicalDeliveryReportRequest = 0x0C0C
const MAPIPhysicalForwardingAddress = 0x0C0D
const MAPIPhysicalForwardingAddressRequested = 0x0C0E
const MAPIPhysicalForwardingProhibited = 0x0C0F
const MAPIPhysicalRenditionAttributes = 0x0C10
const MAPIProofOfDelivery = 0x0C11
const MAPIProofOfDeliveryRequested = 0x0C12
const MAPIRecipientCertificate = 0x0C13
const MAPIRecipientNumberForAdvice = 0x0C14
const MAPIRecipientType = 0x0C15
const MAPIRegisteredMailType = 0x0C16
const MAPIReplyRequested = 0x0C17
const MAPIRequestedDeliveryMethod = 0x0C18
const MAPISenderEntryID = 0x0C19
const MAPISenderName = 0x0C1A
const MAPISupplementaryInfo = 0x0C1B
const MAPITypeOfMtsUser = 0x0C1C
const MAPISenderSearchKey = 0x0C1D
const MAPISenderAddrtype = 0x0C1E
const MAPISenderEmailAddress = 0x0C1F
const MAPICurrentVersion = 0x0E00
const MAPIDeleteAfterSubmit = 0x0E01
const MAPIDisplayBcc = 0x0E02
const MAPIDisplayCc = 0x0E03
const MAPIDisplayTo = 0x0E04
const MAPIParentDisplay = 0x0E05
const MAPIMessageDeliveryTime = 0x0E06
const MAPIMessageFlags = 0x0E07
const MAPIMessageSize = 0x0E08
const MAPIParentEntryID = 0x0E09
const MAPISentmailEntryID = 0x0E0A
const MAPICorrelate = 0x0E0C
const MAPICorrelateMtsID = 0x0E0D
const MAPIDiscreteValues = 0x0E0E
const MAPIResponsibility = 0x0E0F
const MAPISpoolerStatus = 0x0E10
const MAPITransportStatus = 0x0E11
const MAPIMessageRecipients = 0x0E12
const MAPIMessageAttachments = 0x0E13
const MAPISubmitFlags = 0x0E14
const MAPIRecipientStatus = 0x0E15
const MAPITransportKey = 0x0E16
const MAPIMsgStatus = 0x0E17
const MAPIMessageDownloadTime = 0x0E18
const MAPICreationVersion = 0x0E19
const MAPIModifyVersion = 0x0E1A
const MAPIHasattach = 0x0E1B
const MAPIBodyCrc = 0x0E1C
const MAPINormalizedSubject = 0x0E1D
const MAPIRtfInSync = 0x0E1F
const MAPIAttachSize = 0x0E20
const MAPIAttachNum = 0x0E21
const MAPIPreprocess = 0x0E22
const MAPIOriginatingMtaCertificate = 0x0E25
const MAPIProofOfSubmission = 0x0E26
const MAPIEntryID = 0x0FFF
const MAPIObjectType = 0x0FFE
const MAPIIcon = 0x0FFD
const MAPIMiniIcon = 0x0FFC
const MAPIStoreEntryID = 0x0FFB
const MAPIStoreRecordKey = 0x0FFA
const MAPIRecordKey = 0x0FF9
const MAPIMappingSignature = 0x0FF8
const MAPIAccessLevel = 0x0FF7
const MAPIInstanceKey = 0x0FF6
const MAPIRowType = 0x0FF5
const MAPIAccess = 0x0FF4
const MAPIRowID = 0x3000
const MAPIDisplayName = 0x3001
const MAPIAddrtype = 0x3002
const MAPIEmailAddress = 0x3003
const MAPIComment = 0x3004
const MAPIDepth = 0x3005
const MAPIProviderDisplay = 0x3006
const MAPICreationTime = 0x3007
const MAPILastModificationTime = 0x3008
const MAPIResourceFlags = 0x3009
const MAPIProviderDllName = 0x300A
const MAPISearchKey = 0x300B
const MAPIProviderUID = 0x300C
const MAPIProviderOrdinal = 0x300D
const MAPIFormVersion = 0x3301
const MAPIFormClsid = 0x3302
const MAPIFormContactName = 0x3303
const MAPIFormCategory = 0x3304
const MAPIFormCategorySub = 0x3305
const MAPIFormHostMap = 0x3306
const MAPIFormHidden = 0x3307
const MAPIFormDesignerName = 0x3308
const MAPIFormDesignerGuID = 0x3309
const MAPIFormMessageBehavior = 0x330A
const MAPIDefaultStore = 0x3400
const MAPIStoreSupportMask = 0x340D
const MAPIStoreState = 0x340E
const MAPIIpmSubtreeSearchKey = 0x3410
const MAPIIpmOutboxSearchKey = 0x3411
const MAPIIpmWastebasketSearchKey = 0x3412
const MAPIIpmSentmailSearchKey = 0x3413
const MAPIMdbProvider = 0x3414
const MAPIReceiveFolderSettings = 0x3415
const MAPIValidFolderMask = 0x35DF
const MAPIIpmSubtreeEntryID = 0x35E0
const MAPIIpmOutboxEntryID = 0x35E2
const MAPIIpmWastebasketEntryID = 0x35E3
const MAPIIpmSentmailEntryID = 0x35E4
const MAPIViewsEntryID = 0x35E5
const MAPICommonViewsEntryID = 0x35E6
const MAPIFinderEntryID = 0x35E7
const MAPIContainerFlags = 0x3600
const MAPIFolderType = 0x3601
const MAPIContentCount = 0x3602
const MAPIContentUnread = 0x3603
const MAPICreateTemplates = 0x3604
const MAPIDetailsTable = 0x3605
const MAPISearch = 0x3607
const MAPISelectable = 0x3609
const MAPISubfolders = 0x360A
const MAPIStatus = 0x360B
const MAPIAnr = 0x360C
const MAPIContentsSortOrder = 0x360D
const MAPIContainerHierarchy = 0x360E
const MAPIContainerContents = 0x360F
const MAPIFolderAssociatedContents = 0x3610
const MAPIDefCreateDl = 0x3611
const MAPIDefCreateMailuser = 0x3612
const MAPIContainerClass = 0x3613
const MAPIContainerModifyVersion = 0x3614
const MAPIAbProviderID = 0x3615
const MAPIDefaultViewEntryID = 0x3616
const MAPIAssocContentCount = 0x3617
const MAPIAttachmentX400Parameters = 0x3700
const MAPIAttachDataObj = 0x3701
const MAPIAttachEncoding = 0x3702
const MAPIAttachExtension = 0x3703
const MAPIAttachFilename = 0x3704
const MAPIAttachMethod = 0x3705
const MAPIAttachLongFilename = 0x3707
const MAPIAttachPathname = 0x3708
const MAPIAttachRendering = 0x3709
const MAPIAttachTag = 0x370A
const MAPIRenderingPosition = 0x370B
const MAPIAttachTransportName = 0x370C
const MAPIAttachLongPathname = 0x370D
const MAPIAttachMimeTag = 0x370E
const MAPIAttachAdditionalInfo = 0x370F
const MAPIDisplayType = 0x3900
const MAPITemplateID = 0x3902
const MAPIPrimaryCapability = 0x3904
const MAPI7bitDisplayName = 0x39FF
const MAPIAccount = 0x3A00
const MAPIAlternateRecipient = 0x3A01
const MAPICallbackTelephoneNumber = 0x3A02
const MAPIConversionProhibited = 0x3A03
const MAPIDiscloseRecipients = 0x3A04
const MAPIGeneration = 0x3A05
const MAPIGivenName = 0x3A06
const MAPIGovernmentIDNumber = 0x3A07
const MAPIBusinessTelephoneNumber = 0x3A08
const MAPIHomeTelephoneNumber = 0x3A09
const MAPIInitials = 0x3A0A
const MAPIKeyword = 0x3A0B
const MAPILanguage = 0x3A0C
const MAPILocation = 0x3A0D
const MAPIMailPermission = 0x3A0E
const MAPIMhsCommonName = 0x3A0F
const MAPIOrganizationalIDNumber = 0x3A10
const MAPISurname = 0x3A11
const MAPIOriginalEntryID = 0x3A12
const MAPIOriginalDisplayName = 0x3A13
const MAPIOriginalSearchKey = 0x3A14
const MAPIPostalAddress = 0x3A15
const MAPICompanyName = 0x3A16
const MAPITitle = 0x3A17
const MAPIDepartmentName = 0x3A18
const MAPIOfficeLocation = 0x3A19
const MAPIPrimaryTelephoneNumber = 0x3A1A
const MAPIBusiness2TelephoneNumber = 0x3A1B
const MAPIMobileTelephoneNumber = 0x3A1C
const MAPIRadioTelephoneNumber = 0x3A1D
const MAPICarTelephoneNumber = 0x3A1E
const MAPIOtherTelephoneNumber = 0x3A1F
const MAPITransmitableDisplayName = 0x3A20
const MAPIPagerTelephoneNumber = 0x3A21
const MAPIUserCertificate = 0x3A22
const MAPIPrimaryFaxNumber = 0x3A23
const MAPIBusinessFaxNumber = 0x3A24
const MAPIHomeFaxNumber = 0x3A25
const MAPICountry = 0x3A26
const MAPILocality = 0x3A27
const MAPIStateOrProvince = 0x3A28
const MAPIStreetAddress = 0x3A29
const MAPIPostalCode = 0x3A2A
const MAPIPostOfficeBox = 0x3A2B
const MAPITelexNumber = 0x3A2C
const MAPIIsdnNumber = 0x3A2D
const MAPIAssistantTelephoneNumber = 0x3A2E
const MAPIHome2TelephoneNumber = 0x3A2F
const MAPIAssistant = 0x3A30
const MAPISendRichInfo = 0x3A40
const MAPIWeddingAnniversary = 0x3A41
const MAPIBirthday = 0x3A42
const MAPIHobbies = 0x3A43
const MAPIMiddleName = 0x3A44
const MAPIDisplayNamePrefix = 0x3A45
const MAPIProfession = 0x3A46
const MAPIPreferredByName = 0x3A47
const MAPISpouseName = 0x3A48
const MAPIComputerNetworkName = 0x3A49
const MAPICustomerID = 0x3A4A
const MAPITtytddPhoneNumber = 0x3A4B
const MAPIFtpSite = 0x3A4C
const MAPIGender = 0x3A4D
const MAPIManagerName = 0x3A4E
const MAPINickname = 0x3A4F
const MAPIPersonalHomePage = 0x3A50
const MAPIBusinessHomePage = 0x3A51
const MAPIContactVersion = 0x3A52
const MAPIContactEntryids = 0x3A53
const MAPIContactAddrtypes = 0x3A54
const MAPIContactDefaultAddressIndex = 0x3A55
const MAPIContactEmailAddresses = 0x3A56
const MAPICompanyMainPhoneNumber = 0x3A57
const MAPIChildrensNames = 0x3A58
const MAPIHomeAddressCity = 0x3A59
const MAPIHomeAddressCountry = 0x3A5A
const MAPIHomeAddressPostalCode = 0x3A5B
const MAPIHomeAddressStateOrProvince = 0x3A5C
const MAPIHomeAddressStreet = 0x3A5D
const MAPIHomeAddressPostOfficeBox = 0x3A5E
const MAPIOtherAddressCity = 0x3A5F
const MAPIOtherAddressCountry = 0x3A60
const MAPIOtherAddressPostalCode = 0x3A61
const MAPIOtherAddressStateOrProvince = 0x3A62
const MAPIOtherAddressStreet = 0x3A63
const MAPIOtherAddressPostOfficeBox = 0x3A64
const MAPIStoreProviders = 0x3D00
const MAPIAbProviders = 0x3D01
const MAPITransportProviders = 0x3D02
const MAPIDefaultProfile = 0x3D04
const MAPIAbSearchPath = 0x3D05
const MAPIAbDefaultDir = 0x3D06
const MAPIAbDefaultPab = 0x3D07
const MAPIFilteringHooks = 0x3D08
const MAPIServiceName = 0x3D09
const MAPIServiceDllName = 0x3D0A
const MAPIServiceEntryName = 0x3D0B
const MAPIServiceUID = 0x3D0C
const MAPIServiceExtraUids = 0x3D0D
const MAPIServices = 0x3D0E
const MAPIServiceSupportFiles = 0x3D0F
const MAPIServiceDeleteFiles = 0x3D10
const MAPIAbSearchPathUpdate = 0x3D11
const MAPIProfileName = 0x3D12
const MAPIIdentityDisplay = 0x3E00
const MAPIIdentityEntryID = 0x3E01
const MAPIResourceMethods = 0x3E02
const MAPIResourceType = 0x3E03
const MAPIStatusCode = 0x3E04
const MAPIIdentitySearchKey = 0x3E05
const MAPIOwnStoreEntryID = 0x3E06
const MAPIResourcePath = 0x3E07
const MAPIStatusString = 0x3E08
const MAPIX400DeferredDeliveryCancel = 0x3E09
const MAPIHeaderFolderEntryID = 0x3E0A
const MAPIRemoteProgress = 0x3E0B
const MAPIRemoteProgressText = 0x3E0C
const MAPIRemoteValidateOk = 0x3E0D
const MAPIControlFlags = 0x3F00
const MAPIControlStructure = 0x3F01
const MAPIControlType = 0x3F02
const MAPIDeltax = 0x3F03
const MAPIDeltay = 0x3F04
const MAPIXpos = 0x3F05
const MAPIYpos = 0x3F06
const MAPIControlID = 0x3F07
const MAPIInitialDetailsPane = 0x3F08
const MAPIIdSecureMin = 0x67F0
const MAPIIdSecureMax = 0x67FF

module.exports = {
    decodeMapi,
    getTypeSize
}