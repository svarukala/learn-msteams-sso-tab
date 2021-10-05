import * as React from "react";
import { Provider, Flex, Text, Button, Header, List } from "@fluentui/react-northstar";
import { useState, useEffect, useCallback } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";

/**
 * Implementation of the SSO Tab content page
 */
export const SsoTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [ssoToken, setSsoToken] = useState<string>();
    const [msGraphOboToken, setMsGraphOboToken] = useState<string>();
    const [spoOboToken, setSPOOboToken] = useState<string>();
    const [recentMail, setRecentMail] = useState<any[]>();
    const [searchResults, setSearchResults] = useState<any[]>();

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                    setName(decoded!.name);
                    setSsoToken(token);
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.TAB_APP_URI as string]
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
        }
    }, [context]);   

    const exchangeSsoTokenForOboTokenForSPO = async () => {
      const response = await fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?ssoToken=${ssoToken}&tokenFor=spo`);
      const responsePayload = await response.json();
      if (response.ok) {
        setSPOOboToken(responsePayload.access_token);
      } else {
        if (responsePayload!.error === "consent_required") {
          setError("consent_required");
        } else {
          setError("unknown SSO error");
        }
      }
    };
  

  const exchangeSsoTokenForOboToken = useCallback(async () => {
      const response = await fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?ssoToken=${ssoToken}&tokenFor=msg`);
      const responsePayload = await response.json();
      if (response.ok) {
        setMsGraphOboToken(responsePayload.access_token);
      } else {
        if (responsePayload!.error === "consent_required") {
          setError("consent_required");
        } else {
          setError("unknown SSO error");
        }
      }
    }, [ssoToken]);

    /*
    const exchangeSsoTokenForOboTokenForSPO = async () => {
        const response = await fetch(`/exchangeSsoTokenForOboToken/?ssoToken=${ssoToken}&tokenFor=spo`);
        const responsePayload = await response.json();
        if (response.ok) {
          setSPOOboToken(responsePayload.access_token);
        } else {
          if (responsePayload!.error === "consent_required") {
            setError("consent_required");
          } else {
            setError("unknown SSO error");
          }
        }
      };
    

    const exchangeSsoTokenForOboToken = useCallback(async () => {
        const response = await fetch(`/exchangeSsoTokenForOboToken/?ssoToken=${ssoToken}`);
        const responsePayload = await response.json();
        if (response.ok) {
          setMsGraphOboToken(responsePayload.access_token);
        } else {
          if (responsePayload!.error === "consent_required") {
            setError("consent_required");
          } else {
            setError("unknown SSO error");
          }
        }
      }, [ssoToken]);

      */

    const getRecentEmails = useCallback(async () => {
        if (!msGraphOboToken) { return; }
      
        const endpoint = `https://graph.microsoft.com/v1.0/me/messages?$select=receivedDateTime,subject&$orderby=receivedDateTime&$top=10`;
        const requestObject = {
          method: 'GET',
          headers: {
            "authorization": "bearer " + msGraphOboToken
          }
        };
      
        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
      
        if (response.ok) {
          const recentMail = responsePayload.value.map((mail: any) => ({
            key: mail.id,
            header: mail.subject,
            headerMedia: mail.receivedDateTime
          }));
          setRecentMail(recentMail);
        }
      }, [msGraphOboToken]);

    useEffect(() => {
        // if the SSO token is defined...
        if (ssoToken && ssoToken.length > 0) {
          exchangeSsoTokenForOboToken();
        }
      }, [exchangeSsoTokenForOboToken, ssoToken]);      

    useEffect(() => {
        getRecentEmails();
      }, [msGraphOboToken]);

      const getSPOSearchResutls = async () => {
        if (!spoOboToken) { return; }
      
        const endpoint = `https://m365x229910.sharepoint.com/_api/search/query?querytext=%27*%27&selectproperties=%27Author,Path,Title,Url%27&rowlimit=10`;
        const requestObject = {
          method: 'GET',
          headers: {
            "authorization": "bearer " + spoOboToken,
            "accept": "application/json; odata=nometadata"
          }
        };
      
        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
      
        console.log(responsePayload.value);
        if (response.ok) {
            const resultSet = responsePayload.PrimaryQueryResult.RelevantResults.Table.Rows.map((result: any) => ({
              key:result.Cells[8].Value,
              header:result.Cells[0].Value,
              headerMedia:result.Cells[2].Value,
              content:result.Cells[1].Value,
            }));        
        console.log(JSON.stringify(resultSet));
        setSearchResults(resultSet);
      }
    }

      
    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="This is your tab" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={`Hello ${name}`} />
                        </div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                            <Button onClick={() => alert("It worked!")}>A sample button</Button>
                        </div>

                        <div>
                            <Button onClick={() => exchangeSsoTokenForOboTokenForSPO()}>Get SPO Token</Button>
                        </div>
                        <div>
                            <Button onClick={() => getSPOSearchResutls()}>Get SPO Search Results</Button>
                        </div>

                        {searchResults && <div><h3>Your search results:</h3><List items={searchResults} /></div>}
                        <br/>
                        <div>
                            {spoOboToken && <div><Text styles={{
                                                                    wordWrap:"break-word"
                                                                }} content={`OBO Token for SPO: ${spoOboToken}`} /></div>}
                        </div>
                        <br/>
                        <div>
                            {ssoToken && <div><Text styles={{
                                                                wordWrap:"break-word"
                                                            }} content={`SSO Token AKA ID Token: ${ssoToken}`} /></div>}
                        </div>
                        <br/>
                        <div>
                            {msGraphOboToken && <div><Text styles={{
                                                                        wordWrap:"break-word"
                                                                    }} content={`OBO Token for MS Graph: ${msGraphOboToken}`} /></div>}
                        </div>
                        <br/>

                        {recentMail && <div><h3>Your recent emails:</h3><List items={recentMail} /></div>}
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright C0nt0s0" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
