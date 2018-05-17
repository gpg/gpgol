/* dialogs.h

   Resouces IDs for the dialogs.
 */

#ifndef DIALOGS_H
#define DIALOGS_H


#define IDB_LOGO                        0x1060


/* Ids for the extended options dialog.  */
#define IDD_EXT_OPTIONS                 0x4110
#define IDC_T_DEBUG_LOGFILE             0x4120
#define IDC_DEBUG_LOGFILE               0x4130


/* Ids used for the main config dialog.  */
#define IDD_GPG_OPTIONS                 0x5000
#define IDD_ADDIN_OPTIONS               0x5001
#define IDC_TIME_PHRASES                0x5010
#define IDC_ENCRYPT_DEFAULT             0x5020
#define IDC_SIGN_DEFAULT                0x5030
#define IDC_ENCRYPT_WITH_STANDARD_KEY   0x5040
#define IDC_OPENPGP_DEFAULT             0x5050
#define IDC_SMIME_DEFAULT               0x5060
#define IDC_GPG_OPTIONS                 0x5070
#define IDC_ADDIN_OPTIONS               0x5071
#define IDC_BITMAP                      0x5080
#define IDC_VERSION_INFO                0x5090
#define IDC_ENCRYPT_TO                  0x50A0
#define IDC_ENABLE_SMIME                0x50B0
#define IDC_PREVIEW_DECRYPT             0x50C0
#define IDC_PREFER_HTML                 0x50D0
#define IDC_G_GENERAL                   0x50E0
#define IDC_G_SEND                      0x50F0
#define IDC_G_RECV                      0x5100
#define IDC_BODY_AS_ATTACHMENT          0x5110
#define IDC_GPG_CONF                    0x5120
#define IDC_G10CODE_STRING              0x5130
#define IDC_GPG4WIN_STRING              0x5131
#define IDC_START_CERTMAN               0x5132
#define IDC_MIME_UI                     0x5133
#define IDC_INLINE_PGP                  0x5134
#define IDC_AUTORRESOLVE                0x5135
#define IDC_REPLYCRYPT                  0x5136

/* Ids for PNG Images */
#define IDI_ENCRYPT_16_PNG              0x6000
#define IDI_ENCRYPT_48_PNG              0x6010
#define IDI_DECRYPT_16_PNG              0x6020
#define IDI_DECRYPT_48_PNG              0x6030
#define IDI_ENCSIGN_FILE_48_PNG         0x6050
#define IDI_SIGN_48_PNG                 0x6060
#define IDI_VERIFY_48_PNG               0x6070
#define IDI_EMBLEM_WARNING_64_PNG       0x6071
#define IDI_EMBLEM_QUESTION_64_PNG      0x6074
#define IDI_SIGN_ENCRYPT_40_PNG         0x6075
#define IDI_ENCRYPT_20_PNG              0x6076
#define IDI_SIGN_20_PNG                 0x6077
#define IDI_GPGOL_LOCK_ICON             0x6078

/* Status icons */
#define ENCRYPT_ICON_OFFSET             0x10

#define IDI_LEVEL_0                     0x6080
#define IDI_LEVEL_1                     0x6081
#define IDI_LEVEL_2                     0x6082
#define IDI_LEVEL_3                     0x6083
#define IDI_LEVEL_4                     0x6084
#define IDI_LEVEL_0_ENC                 (IDI_LEVEL_0 + ENCRYPT_ICON_OFFSET)
#define IDI_LEVEL_1_ENC                 (IDI_LEVEL_1 + ENCRYPT_ICON_OFFSET)
#define IDI_LEVEL_2_ENC                 (IDI_LEVEL_2 + ENCRYPT_ICON_OFFSET)
#define IDI_LEVEL_3_ENC                 (IDI_LEVEL_3 + ENCRYPT_ICON_OFFSET)
#define IDI_LEVEL_4_ENC                 (IDI_LEVEL_4 + ENCRYPT_ICON_OFFSET)

#endif /*DIALOGS_H*/
