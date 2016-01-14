<?php
require_once dirname(__FILE__).'/src/Google/autoload.php';
@file_put_contents("now-cb.txt","point1");
session_start();

$client = new Google_Client();
$client->setAuthConfigFile('client_secret.json');
$client->setAccessType("offline");
//$client->setRedirectUri('http://' . $_SERVER['HTTP_HOST'] . '/ForAmo/cb.php');
$client->setRedirectUri('http://calgel.ru/esp/cb.php');
$client->addScope("https://www.googleapis.com/auth/drive");
@file_put_contents("now-cb.txt","point2",FILE_APPEND);
if (! isset($_GET['code'])) {
    @file_put_contents("now-cb.txt","point2.1",FILE_APPEND);
    $auth_url = $client->createAuthUrl();
    header('Location: ' . filter_var($auth_url, FILTER_SANITIZE_URL));
} else {
    @file_put_contents("now-cb.txt","point2.2 get=".$_GET['code'],FILE_APPEND);
    $client->authenticate($_GET['code']);
    @file_put_contents("now-cb.txt","point2.3",FILE_APPEND);
    $_SESSION['upload_token'] = $client->getAccessToken();
    @file_put_contents("now-cb.txt","point2.4",FILE_APPEND);
    file_put_contents("token", $_SESSION['upload_token']);
    @file_put_contents("now-cb.txt","point2.5",FILE_APPEND);
    $redirect_uri = 'https://calgelacademy.amocrm.ru/';
    @file_put_contents("now-cb.txt","point2.6".filter_var($redirect_uri, FILTER_SANITIZE_URL),FILE_APPEND);
    header('Location: ' . filter_var($redirect_uri, FILTER_SANITIZE_URL));
}