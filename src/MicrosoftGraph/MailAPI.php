<?php

// tenantID : f8cdef31-a31e-4b4a-93e4-5f571e91255a
// Client ID : 8efe8f13-1d5a-4999-8fb6-95785f4bbd80
// Client Secret : 00a7220b-5e32-429d-9b39-25588a8f1490
namespace MicrosoftGraph;

class MailAPI {

    var $tenantID;
    var $clientID;
    var $clientSecret;
    var $Token;
    var $baseURL;
    var $mailbox;

    function __construct($Mailbox, $sTenantID, $sClientID, $sClientSecret) {
        $this->mailbox = $Mailbox;
        $this->tenantID = $sTenantID;
        $this->clientID = $sClientID;
        $this->clientSecret = $sClientSecret;
        $this->baseURL = 'https://graph.microsoft.com/v1.0/';
        $this->Token = $this->getToken();
    }

    function getToken() {
        
        $oauthRequest = 'client_id=' . $this->clientID . '&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret=' . $this->clientSecret . '&grant_type=client_credentials';
        $reply = $this->POST_Request('https://login.microsoftonline.com/' . $this->tenantID . '/oauth2/v2.0/token', $oauthRequest);
        $reply = json_decode($reply['data']);
        // echo json_encode($reply);
        return $reply->access_token;
    }


    function send_email_files($mail_to_r, $mail_replay, $mail_sub, $mail_body, $files_r){

        
        if (!$this->Token) {
            throw new Exception('No token defined');
        }

        //  Array of email address : 
        foreach ($mail_to_r as $adress) {
            $messageArray['toRecipients'][] = array('emailAddress' => array('address' => $adress));
        }
        // Subject :
        $messageArray['subject'] = $mail_sub;
        // Replay : 
        $messageArray['replyTo'] = array(array('emailAddress' => array('address' => $mail_replay )));
        // Body :
        $messageArray['body'] = array('contentType' => 'HTML', 'content' => $mail_body);
        $messageJSON = json_encode($messageArray);
        $response = $this->POST_Request($this->baseURL . 'users/' . $this->mailbox . '/messages', $messageJSON, array('Content-type: application/json'));

        

        $response = json_decode($response['data']);

        if( isset($response->error) ){

            return array(
                'ERR' => $response->error->message
            );
        }
        
        $messageID = $response->id;

        // Attachements :
        foreach ($files_r as $attachment) {

            $attachment_Name = basename($attachment);
            $attachment_Content = file_get_contents($attachment);
            $attachment_ContentType = mime_content_type($attachment);


            $messageJSON = json_encode(
                                    array(
                                        '@odata.type' => '#microsoft.graph.fileAttachment', 
                                        'name' => $attachment_Name, 
                                        'contentBytes' => base64_encode($attachment_Content), 
                                        'contentType' => $attachment_ContentType, 
                                        'isInline' => false
                                    )
                            );

            $response = $this->POST_Request($this->baseURL . 'users/' . $this->mailbox . '/messages/' . $messageID . '/attachments', $messageJSON, array('Content-type: application/json'));
        }



        $response = $this->POST_Request($this->baseURL . 'users/' . $this->mailbox . '/messages/' . $messageID . '/send', '', array('Content-Length: 0'));
        
        if ($response['code'] == '202') return "SUCCESS";
        return array(
                'ERR' => $response['error']
            );
    }
 
    function POST_Request($URL, $Fields, $Headers = false) {
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_POST, 1);
        if ($Fields) curl_setopt($ch, CURLOPT_POSTFIELDS, $Fields);
        if ($Headers) {
            $Headers[] = 'Authorization: Bearer ' . $this->Token;
            curl_setopt($ch, CURLOPT_HTTPHEADER, $Headers);
        }
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        $responseCode = curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
        curl_close($ch);
        return array('code' => $responseCode, 'data' => $response);
    }

    function GET_Request($URL) {
        $ch = curl_init($URL);
        curl_setopt($ch, CURLOPT_HTTPHEADER, array('Authorization: Bearer ' . $this->Token, 'Content-Type: application/json'));
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $response = curl_exec($ch);
        curl_close($ch);
        return $response;
    }
}
?>