#ifndef INC_GPGMESSAGE_H
#define INC_GPGMESSAGE_H

enum pgp_message_t
{
    PGP_NONE=0,
    PGP_MSG=1,
    PGP_SIG=2,
    PGP_CSIG=4,
    PGP_KEY=8,
};

class CGPGMessage
{
public:
    CGPGMessage (LPMESSAGE pMsg);
    ~CGPGMessage ();

private:
    LPMESSAGE pMessage;
    LPSTREAM pStreamBody;
    LPSPropValue sPropVal;
    SPropValue sProp;
    HRESULT hr;

public:
    int hasAttach (void);
    int hasGPGData (void);

    int isUnset (void);

    int getGPGType (void);
    int getBody (char ** pBody);
    int setBody (char * pBody);

    int msgFlags (void);
};

#endif /*INC_GPGMESSAGE_H*/
