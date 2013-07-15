/* dialogs.h

   Resouces IDs for the dialogs.
 */

#ifndef DIALOGS_H
#define DIALOGS_H


/* Ids used for bitmaps.  There is some magic in the identifiers: In
   the code we only use values like 3000, the icon code then uses the
   current system pixel size to the Id and tries to load this one.
   The mask is always the next id.  */
#define IDB_ENCRYPT                     3000
#define IDB_ENCRYPT_16                  3016
#define IDB_ENCRYPT_16M                 3017
#define IDB_ENCRYPT_32                  3032
#define IDB_ENCRYPT_32M                 3033
#define IDB_ENCRYPT_64                  3064
#define IDB_ENCRYPT_64M                 3065

#define IDB_SIGN                        3100
#define IDB_SIGN_16                     3116
#define IDB_SIGN_16M                    3117
#define IDB_SIGN_32                     3132
#define IDB_SIGN_32M                    3133
#define IDB_SIGN_64                     3164
#define IDB_SIGN_64M                    3165

#define IDB_KEY_MANAGER                 3200
#define IDB_KEY_MANAGER_16              3216
#define IDB_KEY_MANAGER_16M             3217
#define IDB_KEY_MANAGER_32              3232
#define IDB_KEY_MANAGER_32M             3233
#define IDB_KEY_MANAGER_64              3264
#define IDB_KEY_MANAGER_64M             3265

#define IDB_DECRYPT                     3300
#define IDB_DECRYPT_16                  3316
#define IDB_DECRYPT_16M                 3317
#define IDB_DECRYPT_32                  3332
#define IDB_DECRYPT_32M                 3333
#define IDB_DECRYPT_64                  3364
#define IDB_DECRYPT_64M                 3365

#define IDB_VERIFY                      3400
#define IDB_VERIFY_16                   3416
#define IDB_VERIFY_16M                  3417
#define IDB_VERIFY_32                   3432
#define IDB_VERIFY_32M                  3433
#define IDB_VERIFY_64                   3464
#define IDB_VERIFY_64M                  3465

#define IDB_DECRYPT_VERIFY              3500
#define IDB_DECRYPT_VERIFY_16           3516
#define IDB_DECRYPT_VERIFY_16M          3517
#define IDB_DECRYPT_VERIFY_32           3532
#define IDB_DECRYPT_VERIFY_32M          3533
#define IDB_DECRYPT_VERIFY_64           3564
#define IDB_DECRYPT_VERIFY_64M          3565


/* Special */
#define IDB_BANNER                      3900  /* The g10 Code logo.  */


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
#define IDC_G10CODE_STRING              4028


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

/* Ids for PNG Images */
#define IDI_ENCRYPT_16_PNG              6000
#define IDI_ENCRYPT_48_PNG              6001
#define IDI_DECRYPT_16_PNG              6010
#define IDI_DECRYPT_48_PNG              6011
#define IDI_KEY_MANAGER_64_PNG          6020
#define IDI_ENCSIGN_FILE_48_PNG         6030

#endif /*DIALOGS_H*/
