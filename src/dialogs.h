/* dialogs.h

   Resouces IDs for the dialogs.
 */

#ifndef DIALOGS_H
#define DIALOGS_H


/* Ids used for bitmaps.  */
#define IDB_DECRYPT                     3001  /* Not used.  */
#define IDB_ENCRYPT                     3002
#define IDB_SIGN                        3003
#define IDB_ADD_KEYS                    3004
#define IDB_KEY_MANAGER                 3005
#define IDB_BANNER                      3006  /* The g10 Code logo.  */
#define IDB_BANNER_HI                   3007  /* Not used.  */
#define IDB_CRYPTO_STATE                3008  
#define IDB_PROTO_AUTO                  3009
#define IDB_PROTO_PGPMIME               3010
#define IDB_PROTO_SMIME                 3011


/* Ids used for the main config dialog.  */
#define IDD_GPG_OPTIONS                 4001
#define IDC_TIME_PHRASES                4010
#define IDC_ENCRYPT_DEFAULT             4011
#define IDC_SIGN_DEFAULT                4012
#define IDC_ENCRYPT_WITH_STANDARD_KEY   4013
#define IDC_OPENPGP_DEFAULT             4014
#define IDC_SMIME_DEFAULT               4015
#define IDC_GPG_OPTIONS                 4016
#define IDC_BITMAP                      4017
#define IDC_VERSION_INFO                4018
#define IDC_ENCRYPT_TO                  4019
#define IDC_ENABLE_SMIME                4020
#define IDC_PREVIEW_DECRYPT             4021
#define IDC_PREFER_HTML                 4022
#define IDC_G_GENERAL                   4023
#define IDC_G_SEND                      4024
#define IDC_G_RECV                      4025
#define IDC_BODY_AS_ATTACHMENT          4026
#define IDC_GPG_CONF                    4027


/* Ids for the extended options dialog.  */
#define IDD_EXT_OPTIONS                 4101
#define IDC_T_DEBUG_LOGFILE             4113
#define IDC_DEBUG_LOGFILE               4114


/* Ids for the recipient selection dialog.  */
#define IDD_ENC                         4201
#define IDC_ENC_RSET1                   4210
#define IDC_ENC_RSET2_T                 4211
#define IDC_ENC_RSET2                   4212
#define IDC_ENC_NOTFOUND_T              4213
#define IDC_ENC_NOTFOUND                4214


/* Ids for the two decryption dialogs.  */
#define IDD_DEC                         4301
#define IDD_DECEXT                      4302
#define IDC_DEC_KEYLIST                 4310
#define IDC_DEC_HINT                    4311
#define IDC_DEC_PASSINF                 4312
#define IDC_DEC_PASS                    4313
#define IDC_DEC_HIDE                    4314
#define IDC_DECEXT_RSET_T               4320
#define IDC_DECEXT_RSET                 4321
#define IDC_DECEXT_KEYLIST              4322
#define IDC_DECEXT_HINT                 4323
#define IDC_DECEXT_PASSINF              4324
#define IDC_DECEXT_PASS                 4325
#define IDC_DECEXT_HIDE                 4326


/* Ids for the verification dialog.  */
#define IDD_VRY                         4401
#define IDC_VRY_TIME_T                  4410
#define IDC_VRY_TIME                    4411
#define IDC_VRY_PKALGO_T                4412
#define IDC_VRY_PKALGO                  4413
#define IDC_VRY_KEYID_T                 4414
#define IDC_VRY_KEYID                   4415
#define IDC_VRY_STATUS                  4416
#define IDC_VRY_ISSUER_T                4417
#define IDC_VRY_ISSUER                  4418
#define IDC_VRY_AKALIST_T               4419
#define IDC_VRY_AKALIST                 4420
#define IDC_VRY_HINT                    4421



#endif /*DIALOGS_H*/
