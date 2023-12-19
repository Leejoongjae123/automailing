import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import datetime
import time
import random
import json
import requests
from bs4 import BeautifulSoup
import pprint
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def send_email(totalNews,totalArticles,spreadData):
    # 네이버 SMTP 서버 정보
    smtp_server = 'smtp.naver.com'
    smtp_port = 587  # TLS 포트

    # 보내는 사람 이메일 주소와 비밀번호
    sender_email = '1jaeyeon@naver.com'
    sender_password = 'ckdduvie!'

    # 받는 사람 이메일 주소
    # receiver_email = 'jaeyeon.won@chdpharm.com'
    receiver_email = spreadData['메일']

    # 이메일 제목과 본문
    timeNow = datetime.datetime.now().strftime("%m월%d일")
    email_subject = '[DM팀 뉴스레터] {} AI Article'.format(timeNow)
    newsBody=""
    for indexNews,totalNew in enumerate(totalNews):
        newsContents='<p style="margin: 0;font-weight:bold;">{}){}</p>'.format(indexNews+1,totalNew['name'])
        for index,news in enumerate(totalNew['news']):
            newsContents=newsContents+'<p style="margin: 0; padding: 10px 0px 0px 10px;font-size: 14px"><a href={}>{}</a><span style="margin: 0; padding: 10px 0px 0px 10px;font-size: 14px">{}</span></p>'.format(news['url'],news['title'],news['regiDate'])

        newsBody = newsBody + '''
                    <tr>
                        <td style="padding: 0px 40px 20px 40px; font-family: sans-serif; font-size: 16px; line-height: 20px; color: #555555; text-align: left; font-weight:normal;">
                            {}
                        </td>
                    </tr>        
                '''.format(newsContents)

    print("newsBody:",newsBody,"/ newsBody_TYPE:",type(newsBody),len(newsBody))


    paperBody=""
    print("totalArticles:",totalArticles,"/ totalArticles_TYPE:",type(totalArticles),len(totalArticles))
    for indexTotal,totalArticle in enumerate(totalArticles):
        paperContents=""
        for index,paper in enumerate(totalArticle['research']):
            paperContents=paperContents+'''
            <a href='{}'>{}</a><br>
            <p style="margin: 0;font-size: 14px">{}</p>
            '''.format(paper['url'],paper['title'],paper['paper'])
        paperBody = paperBody + '''
                    <tr>
                        <td style="padding: 0px 40px 40px 40px;">
                            <table width="100%" cellspacing="0" cellpadding="0" border="0" bgcolor="#ffffff" style="border:1px solid #dddddd; border-left:3px solid #0159B7;">
                                <tr>
                                    <td align="" style="padding: 20px 20px 0px 20px; font-family: 'Montserrat', sans-serif; font-size: 14px; line-height: 20px; color: #555555; text-align: left; font-weight:normal;">
                                        <h3 style="margin:0;">
                                            {}) {}
                                        </h3>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="" style="padding: 20px 20px 20px 20px; font-family: sans-serif; font-size: 16px; line-height: 20px; color: #555555; text-align: left; font-weight:normal;">
                                        <p style="margin:0;">
                                            {}
                                        </p>
                                    </td>
                                </tr>
                            </table>

                        </td>
                    </tr>                
                    '''.format(indexTotal+1,totalArticle['name'], paperContents)

    print("paperContents:",paperContents,"/ paperContents_TYPE:",type(paperContents),len(paperContents))




    email_body = '''
    <!DOCTYPE html>
    <html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
    <head>
        <meta charset="utf-8"> <!-- utf-8 works for most cases -->
        <meta name="viewport" content="width=device-width"> <!-- Forcing initial-scale shouldn't be necessary -->
        <meta http-equiv="X-UA-Compatible" content="IE=edge"> <!-- Use the latest (edge) version of IE rendering engine -->
        <meta name="x-apple-disable-message-reformatting">  <!-- Disable auto-scale in iOS 10 Mail entirely -->
        <title>Event - [Plain HTML]</title> <!-- The title tag shows in email notifications, like Android 4.4. -->

        <!-- Web Font / @font-face : BEGIN -->
        <!-- NOTE: If web fonts are not required, lines 10 - 27 can be safely removed. -->

        <!-- Desktop Outlook chokes on web font references and defaults to Times New Roman, so we force a safe fallback font. -->
        <!--[if mso]>
            <style>
                * {
                    font-family: Arial, sans-serif !important;
                }
            </style>
        <![endif]-->

        <!-- All other clients get the webfont reference; some will render the font and others will silently fail to the fallbacks. More on that here: http://stylecampaign.com/blog/2015/02/webfont-support-in-email/ -->
        <!--[if !mso]><!-->
            <link href="https://fonts.googleapis.com/css?family=Montserrat:300,500" rel="stylesheet">
        <!--<![endif]-->

        <!-- Web Font / @font-face : END -->

        <!-- CSS Reset -->
        <style>

            /* What it does: Remove spaces around the email design added by some email clients. */
            /* Beware: It can remove the padding / margin and add a background color to the compose a reply window. */
            html,
            body {
                margin: 0 auto !important;
                padding: 0 !important;
                height: 100% !important;
                width: 100% !important;
            }

            /* What it does: Stops email clients resizing small text. */
            * {
                -ms-text-size-adjust: 100%;
                -webkit-text-size-adjust: 100%;
            }

            /* What it does: Centers email on Android 4.4 */
            div[style*="margin: 16px 0"] {
                margin:0 !important;
            }

            /* What it does: Stops Outlook from adding extra spacing to tables. */
            table,
            td {
                mso-table-lspace: 0pt !important;
                mso-table-rspace: 0pt !important;
            }

            /* What it does: Fixes webkit padding issue. Fix for Yahoo mail table alignment bug. Applies table-layout to the first 2 tables then removes for anything nested deeper. */
            table {
                border-spacing: 0 !important;
                border-collapse: collapse !important;
                table-layout: fixed !important;
                margin: 0 auto !important;
            }
            table table table {
                table-layout: auto;
            }

            /* What it does: Uses a better rendering method when resizing images in IE. */
            img {
                -ms-interpolation-mode:bicubic;
            }

            /* What it does: A work-around for email clients meddling in triggered links. */
            *[x-apple-data-detectors],	/* iOS */
            .x-gmail-data-detectors, 	/* Gmail */
            .x-gmail-data-detectors *,
            .aBn {
                border-bottom: 0 !important;
                cursor: default !important;
                color: inherit !important;
                text-decoration: none !important;
                font-size: inherit !important;
                font-family: inherit !important;
                font-weight: inherit !important;
                line-height: inherit !important;
            }

            /* What it does: Prevents Gmail from displaying an download button on large, non-linked images. */
            .a6S {
                display: none !important;
                opacity: 0.01 !important;
            }
            /* If the above doesn't work, add a .g-img class to any image in question. */
            img.g-img + div {
                display:none !important;
               }

            /* What it does: Prevents underlining the button text in Windows 10 */
            .button-link {
                text-decoration: none !important;
            }

            /* What it does: Removes right gutter in Gmail iOS app: https://github.com/TedGoas/Cerberus/issues/89  */
            /* Create one of these media queries for each additional viewport size you'd like to fix */
            /* Thanks to Eric Lepetit @ericlepetitsf) for help troubleshooting */
            @media only screen and (min-device-width: 375px) and (max-device-width: 413px) { /* iPhone 6 and 6+ */
                .email-container {
                    min-width: 375px !important;
                }
            }

        </style>

        <!-- Progressive Enhancements -->
        <style>

            /* What it does: Hover styles for buttons */
            .button-td,
            .button-a {
                transition: all 100ms ease-in;
            }
            .button-td:hover,
            .button-a:hover {
                background: #000000 !important;
                border-color: #000000 !important;
                color: white !important;
            }

            /* Media Queries */
            @media screen and (max-width: 480px) {

                /* What it does: Forces elements to resize to the full width of their container. Useful for resizing images beyond their max-width. */
                .fluid {
                    width: 100% !important;
                    max-width: 100% !important;
                    height: auto !important;
                    margin-left: auto !important;
                    margin-right: auto !important;
                }

                /* What it does: Forces table cells into full-width rows. */
                .stack-column,
                .stack-column-center {
                    display: block !important;
                    width: 100% !important;
                    max-width: 100% !important;
                    direction: ltr !important;
                }
                /* And center justify these ones. */
                .stack-column-center {
                    text-align: center !important;
                }

                /* What it does: Generic utility class for centering. Useful for images, buttons, and nested tables. */
                .center-on-narrow {
                    text-align: center !important;
                    display: block !important;
                    margin-left: auto !important;
                    margin-right: auto !important;
                    float: none !important;
                }
                table.center-on-narrow {
                    display: inline-block !important;
                }

                /* What it does: Adjust typography on small screens to improve readability */
                .email-container p {
                    font-size: 17px !important;
                    line-height: 22px !important;
                }
            }

        </style>

        <!-- What it does: Makes background images in 72ppi Outlook render at correct size. -->
        <!--[if gte mso 9]>
        <xml>
            <o:OfficeDocumentSettings>
                <o:AllowPNG/>
                <o:PixelsPerInch>96</o:PixelsPerInch>
            </o:OfficeDocumentSettings>
        </xml>
        <![endif]-->

    </head>
    ''' + f'''
    <body width="100%" bgcolor="#F1F1F1" style="margin: 0; mso-line-height-rule: exactly;">
        <center style="width: 100%; background: #F1F1F1; text-align: left;">

            <!-- Visually Hidden Preheader Text : BEGIN -->
            <div style="display:none;font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;">
                (Optional) This text will appear in the inbox preview, but not the email body. It can be used to supplement the email subject line or even summarize the email's contents. Extended text preheaders (~490 characters) seems like a better UX for anyone using a screenreader or voice-command apps like Siri to dictate the contents of an email. If this text is not included, email clients will automatically populate it using the text (including image alt text) at the start of the email's body.
            </div>
            <!-- Visually Hidden Preheader Text : END -->

            <!--
                Set the email width. Defined in two places:
                1. max-width for all clients except Desktop Windows Outlook, allowing the email to squish on narrow but never go wider than 680px.
                2. MSO tags for Desktop Windows Outlook enforce a 680px width.
                Note: The Fluid and Responsive templates have a different width (600px). The hybrid grid is more "fragile", and I've found that 680px is a good width. Change with caution.
            -->
            <div style="max-width: 680px; margin: auto;" class="email-container">
                <!--[if mso]>
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="680" align="center">
                <tr>
                <td>
                <![endif]-->

                <!-- Email Body : BEGIN -->
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" align="center" width="100%" style="max-width: 680px;" class="email-container">


                    <!-- HEADER : BEGIN -->
                    <tr>
                        <td bgcolor="#333333">

                        </td>
                    </tr>
                    <!-- HEADER : END -->

                    <!-- HERO : BEGIN -->
                    <tr>
                        <!-- Bulletproof Background Images c/o https://backgrounds.cm -->
                        <td background="background.png" bgcolor="#222222" align="center" valign="top" style="text-align: center; background-position: center center !important; background-size: cover !important;">
                            <!--[if gte mso 9]>
                            <v:rect xmlns:v="urn:schemas-microsoft-com:vml" fill="true" stroke="false" style="width:680px; height:380px; background-position: center center !important;">
                            <v:fill type="tile" src="background.png" color="#222222" />
                            <v:textbox inset="0,0,0,0">
                            <![endif]-->
                            <div>
                                <!--[if mso]>
                                <table role="presentation" border="0" cellspacing="0" cellpadding="0" align="center" width="500">
                                <tr>
                                <td align="center" valign="middle" width="500">
                                <![endif]-->
                                <table role="presentation" border="0" cellpadding="0" cellspacing="0" align="center" width="100%" style="max-width:500px; margin: auto;">

                                    <tr>
                                        <td height="20" style="font-size:20px; line-height:20px;">&nbsp;</td>
                                    </tr>

                                    <tr>
                                      <td align="center" valign="middle">

                                      <table>
                                         <tr>
                                             <td valign="top" style="text-align: center; padding: 60px 0 10px 20px;">

                                                 <h1 style="margin: 0; font-family: 'Montserrat', sans-serif; font-size: 30px; line-height: 36px; color: #ffffff; font-weight: bold;">{timeNow} 당뇨 뉴스레터</h1>
                                             </td>
                                         </tr>
                                         <tr>
                                             <td valign="top" style="text-align: center; padding: 10px 20px 15px 20px; font-family: sans-serif; font-size: 18px; line-height: 20px; color: #ffffff;">
                                                 <p style="margin: 0;">AI Article(네이버 뉴스 및 논문 정보 자동 송신)</p>
                                             </td>
                                         </tr>
                                         <tr>
                                             <td valign="top" align="center" style="text-align: center; padding: 15px 0px 60px 0px;">



                                             </td>
                                         </tr> 
                                      </table>

                                      </td>
                                    </tr>

                                    <tr>
                                        <td height="20" style="font-size:20px; line-height:20px;">&nbsp;</td>
                                    </tr>

                                </table>
                                <!--[if mso]>
                                </td>
                                </tr>
                                </table>
                                <![endif]-->
                            </div>
                            <!--[if gte mso 9]>
                            </v:textbox>
                            </v:rect>
                            <![endif]-->
                        </td>
                    </tr>
                    <!-- HERO : END -->

                    <!-- INTRO : BEGIN -->
                    <tr>
                        <td bgcolor="#ffffff">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="padding: 40px 40px 20px 40px; font-family: sans-serif; font-size: 18px; line-height: 20px; color: #555555; text-align: left; font-weight:bold;">
                                        <p style="margin: 0;">안녕하세요.</p><br>
                                        <p style="margin: 0;">DM팀 뉴스 레터 전달 드립니다.</p><br>
                                        <p style="margin: 0;">Opening Call , DM 품목 디테일 콜에 활용 부탁 드립니다.</p>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 20px 20px 20px 40px; text-align: left;">
                                        <img class="NHN_MAIL_IMAGE" src="https://aiarticleimage.s3.ap-northeast-2.amazonaws.com/NAVER.png" alt="NAVER.png" data-image-original-width="80" data-image-ratio="1.01" data-image-scale="SCALE_CUSTOM" data-image-border="false width="60" height="60"">
                                    </td>
                                </tr>

                                <tr>
                                    <td style="padding: 0px 40px 20px 40px; font-family: sans-serif; font-size: 16px; line-height: 20px; color: #555555; text-align: left; font-weight:bold;">
                                        <p style="margin: 0;">최근 6개월 간 네이버 뉴스</p>
                                    </td>
                                </tr>
                                {newsBody}
                                



                                <tr>
                                    <td align="left" style="padding: 0px 40px 40px 40px;">



                                    </td>
                                </tr>

                            </table>
                        </td>
                    </tr>
                    <!-- INTRO : END -->

                    <!-- AGENDA : BEGIN -->
                    <tr>
                        <td bgcolor="#f7fafc">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="padding: 20px 20px 20px 40px; text-align: left;">
                                        <img class="NHN_MAIL_IMAGE" src="https://aiarticleimage.s3.ap-northeast-2.amazonaws.com/pubmed.png" alt="NAVER.png" data-image-original-width="80" data-image-ratio="1.01" data-image-scale="SCALE_CUSTOM" data-image-border="false width="120" height="60"">
                                    </td>
                                </tr>

                                <tr>
                                    <td style="padding: 0px 40px 35px 40px; font-family: sans-serif; font-size: 18px; line-height: 20px; color: #555555; text-align: left; font-weight:normal;">
                                        <p style="font-family: sans-serif; font-size: 16px; line-height: 20px; color: #555555; text-align: left; font-weight:bold;">최근 6개월 간 PUBMED 논문</p>
                                    </td>
                                </tr>
                                {paperBody}

                            </table>

                        </td>
                    </tr>
                    <!-- AGENDA : END -->



                    <!-- CTA : BEGIN -->
                    <tr>
                        <td bgcolor="#0159B7">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="padding: 40px 40px 5px 40px; text-align: center;">
                                        <h1 style="margin: 0; font-family: 'Montserrat', sans-serif; font-size: 20px; line-height: 24px; color: #ffffff; font-weight: bold;">더 많은 의학 논문 정보를 원하세요?</h1>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 0px 40px 20px 40px; font-family: sans-serif; font-size: 17px; line-height: 23px; color: #C9E4F3; text-align: center; font-weight:normal;">
                                        <p style="margin: 0;">전 세계 다양한 논문을 찾아보세요!</p>
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="middle" align="center" style="text-align: center; padding: 0px 20px 40px 20px;">

                                        <!-- Button : BEGIN -->
                                        <center>
                                        <table role="presentation" align="center" cellspacing="0" cellpadding="0" border="0" class="center-on-narrow">
                                            <tr>
                                                <td style="border-radius: 50px; background: #ffffff; text-align: center;" class="button-td">
                                                    <a href="https://pubmed.ncbi.nlm.nih.gov/" style="background: #ffffff; border: 15px solid #ffffff; font-family: 'Montserrat', sans-serif; font-size: 14px; line-height: 1.1; text-align: center; text-decoration: none; display: block; border-radius: 50px; font-weight: bold;" class="button-a">
                                                        <span style="color:#0159B7;" class="button-link">&nbsp;&nbsp;&nbsp;&nbsp;PUBMED 바로가기&nbsp;&nbsp;&nbsp;&nbsp;</span>
                                                    </a>
                                                </td>
                                            </tr>
                                        </table>
                                        </center>
                                        <!-- Button : END -->

                                    </td>
                                </tr>

                            </table>
                        </td>
                    </tr>
                    <!-- CTA : END -->



                    <!-- FOOTER : BEGIN -->
                    <tr>
                        <td bgcolor="#ffffff">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="padding: 40px 40px 10px 40px; font-family: sans-serif; font-size: 12px; line-height: 18px; color: #666666; text-align: center; font-weight:normal;">
                                        <p style="margin: 0;">POWERED BY AURAWORKS</p>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 0px 40px 10px 40px; font-family: sans-serif; font-size: 12px; line-height: 18px; color: #666666; text-align: center; font-weight:normal;">
                                        <p style="margin: 0;">이 이메일은 <a style="font-weight:bold" href="https://docs.google.com/spreadsheets/d/1UTuXepWAxTeeZUMKGDI8Gd8U1cXYyEQKLLXUUumZma4/edit#gid=0">구글 스프레드 시트</a>에 근거하여 송신됩니다.</p>
                                    </td>
                                </tr>

                                <tr>
                                    <td style="padding: 0px 40px 40px 40px; font-family: sans-serif; font-size: 12px; line-height: 18px; color: #666666; text-align: center; font-weight:normal;">
                                        <p style="margin: 0;">Copyright &copy; 2023-2024 <b>종근당</b>, All Rights Reserved.</p>
                                    </td>
                                </tr>

                            </table>
                        </td>
                    </tr>
                    <!-- FOOTER : END -->

                </table>
                <!-- Email Body : END -->

                <!--[if mso]>
                </td>
                </tr>
                </table>
                <![endif]-->
            </div>

        </center>
    </body>
    </html>
    '''


    # 이메일 메시지 생성
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = email_subject

    # 이메일 본문 추가
    message.attach(MIMEText(email_body, 'html'))

    # SMTP 서버 연결
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # TLS 암호화 시작

    # 로그인
    server.login(sender_email, sender_password)

    # 이메일 발송
    server.sendmail(sender_email, receiver_email, message.as_string())

    # SMTP 서버 연결 종료
    server.quit()

    print("이메일이 성공적으로 보내졌습니다.")

def GetDate(relative_time):
    current_time = datetime.datetime.now()
    if relative_time.find('주')>=0:
        weeks_ago = int(re.findall(r'\d+', relative_time)[0])
        print("weeks_ago:",weeks_ago,"/ weeks_ago_TYPE:",type(weeks_ago))
        current_time -= datetime.timedelta(weeks=weeks_ago)
        current_time=current_time.strftime("%Y.%m.%d")
    elif relative_time.find('일') >= 0:
        days_ago = int(re.findall(r'\d+', relative_time)[0])
        print("days_ago:",days_ago,"/ days_ago_TYPE:",type(days_ago))
        current_time -= datetime.timedelta(days=days_ago)
        current_time = current_time.strftime("%Y.%m.%d")
    elif relative_time.find('분') >= 0:
        minutes_ago = int(re.findall(r'\d+', relative_time)[0])
        print("minutes_ago:",minutes_ago,"/ minutes_ago_TYPE:",type(minutes_ago))
        current_time -= datetime.timedelta(minutes=minutes_ago)
        current_time = current_time.strftime("%Y.%m.%d")
    elif relative_time.find('시간') >= 0:
        hours_ago = int(re.findall(r'\d+', relative_time)[0])
        print("hours_ago:",hours_ago,"/ hours_ago_TYPE:",type(hours_ago))
        current_time -= datetime.timedelta(hours=hours_ago)
        current_time = current_time.strftime("%Y.%m.%d")
    else:
        current_time=relative_time
    print("계산된 날짜:", current_time)
    return current_time
# 이메일 보내기 함수 호출
def GetArticlesPubMed(inputData):
    cookies = {
        'pm-csrf': 'uqWrVoHlodts5cDcnaywmLuKCkzpxQbJ7HnXO3ScL8R2WqriH5SFourkkTj4p3oM',
        'ncbi_sid': '72B77A10531DD433_1347SID',
        'pm-ps': '100',
        'pm-sessionid': 'tpx71l5zw1w6s0bgmza2ozqjah5mraah',
        '_gid': 'GA1.2.625829407.1699243934',
        'pm-sb': 'date',
        'pm-sid': 'PL2hZj-FI6dIQDyD7V2g2w:bbe66251c21403c67dad93be81e2304a',
        'pm-adjnav-sid': 'K_HSWkhKnfBnuXgZpgnBKg:bbe66251c21403c67dad93be81e2304a',
        'pm-iosp': '',
        '_ga_DP2X732JSX': 'GS1.1.1699247400.21.1.1699248069.0.0.0',
        '_ga': 'GA1.1.492663109.1697767439',
        'ncbi_pinger': 'N4IgDgTgpgbg+mAFgSwCYgFwgMwGECsAYgAykCM+AbMYWQCwBC+pp2AgsQOwCiAIgJx06+MrwB0ZMQFs42EABoQAVwB2AGwD2AQ1QqoADwAumUACZM4JQCMpUdIrlYw12/ZB0LAZyhaIAY0RoTyU1Y0V8CwUQMjILWMVTYgs8IhYKalpGZhZ2Lj5BYVEJaVko01inFzsMZxtq718AoJDDDAA5AHk27jLzSrrUMRU/K2QhtSkh5EQxAHMNGDL+OP5EqOwkrDJiAA4khwro1f2cPujiOhPsRxAAMy01b3WPLEMIJSh1nciHZaxsYScKKXCz4fCcZaKOg3YhibD8OHAl7KdTaXQGMLuCJYE74G47ThA8LI/jYeIgKgWXFAnFRSiHMicbCUOksrANfyIEAAX25QA',
    }

    headers = {
        'authority': 'pubmed.ncbi.nlm.nih.gov',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'cache-control': 'max-age=0',
        # 'cookie': 'pm-csrf=uqWrVoHlodts5cDcnaywmLuKCkzpxQbJ7HnXO3ScL8R2WqriH5SFourkkTj4p3oM; ncbi_sid=72B77A10531DD433_1347SID; pm-ps=100; pm-sessionid=tpx71l5zw1w6s0bgmza2ozqjah5mraah; _gid=GA1.2.625829407.1699243934; pm-sb=date; pm-sid=PL2hZj-FI6dIQDyD7V2g2w:bbe66251c21403c67dad93be81e2304a; pm-adjnav-sid=K_HSWkhKnfBnuXgZpgnBKg:bbe66251c21403c67dad93be81e2304a; pm-iosp=; _ga_DP2X732JSX=GS1.1.1699247400.21.1.1699248069.0.0.0; _ga=GA1.1.492663109.1697767439; ncbi_pinger=N4IgDgTgpgbg+mAFgSwCYgFwgMwGECsAYgAykCM+AbMYWQCwBC+pp2AgsQOwCiAIgJx06+MrwB0ZMQFs42EABoQAVwB2AGwD2AQ1QqoADwAumUACZM4JQCMpUdIrlYw12/ZB0LAZyhaIAY0RoTyU1Y0V8CwUQMjILWMVTYgs8IhYKalpGZhZ2Lj5BYVEJaVko01inFzsMZxtq718AoJDDDAA5AHk27jLzSrrUMRU/K2QhtSkh5EQxAHMNGDL+OP5EqOwkrDJiAA4khwro1f2cPujiOhPsRxAAMy01b3WPLEMIJSh1nciHZaxsYScKKXCz4fCcZaKOg3YhibD8OHAl7KdTaXQGMLuCJYE74G47ThA8LI/jYeIgKgWXFAnFRSiHMicbCUOksrANfyIEAAX25QA',
        'referer': 'https://pubmed.ncbi.nlm.nih.gov/advanced/',
        'sec-ch-ua': '"Chromium";v="118", "Google Chrome";v="118", "Not=A?Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36',
    }

    params = {
        'term': inputData['keyword'],
        'sort': 'date',
    }

    response = requests.get('https://pubmed.ncbi.nlm.nih.gov/', params=params, cookies=cookies, headers=headers)


    soup = BeautifulSoup(response.text, 'lxml')
    # print(soup.prettify())
    articles=soup.find_all("article",attrs={'class':'full-docsum'})
    resultList=[]
    for article in articles:
        try:
            title=article.find('a',attrs={'class':'docsum-title'}).get_text().strip()
        except:
            title=""
        # print("title:",title)
        try:
            url='https://pubmed.ncbi.nlm.nih.gov/'+article.find('a',attrs={'class':'docsum-title'})['href']
        except:
            url=""
        # print("url:",url)
        try:
            paper=article.find('span',attrs={'class':'docsum-journal-citation full-journal-citation'}).get_text()
            positionFr=paper.find("doi:")
            paper=paper[:positionFr]
        except:
            paper=""

        # print("paper:",paper)

        data={'title':title,'url':url,'paper':paper}
        resultList.append(data)

    resultList=resultList[:3]
    inputData.update({'research':resultList})
    pprint.pprint(inputData)

    return inputData

def GetNews(inputData):
    count=0
    resultList=[]
    cookies = {
        'NNB': 'EOQEJJQUFMYWK',
        'nx_ssl': '2',
        'ASID': '3d6968c30000018b481054af00008e99',
        '_ga': 'GA1.2.1867842095.1698124262',
        '_fbp': 'fb.1.1698124262166.1924394679',
        '_tt_enable_cookie': '1',
        '_ttp': 'BEmRKkFpqC2qDqQF7ScXQ2-Dz-t',
        '_ga_4BKHBFKFK0': 'GS1.1.1698124262.1.1.1698124263.59.0.0',
        'nid_inf': '-1418340302',
        'NID_AUT': '3SLnksOBVKphncVoHj0APDbP5U/ED2YpQYleT1yokMIn+HQL9NuWvfzlBs8Vg6FE',
        'NID_JKL': 'UWjHrUZIdvQ8BXw7yulMx+BHWLqeT5p9oA1MkpzzkAw=',
        '_naver_usersession_': 'UQOaZKGIfz29lcXjqQmiSA==',
        'NID_SES': 'AAABpCj5RNSPm+yCjzkae5bDyEbeQ5NrJ1dEqogMrSDUDKoyonJClHDOWMCc3KZL/ANrcNVFE+tRQm/JmswWWNzMPEcBZdWtRzlyrQK8bNWTtco1LekPfTTR6VBlFJxBw/4596ANGrhHafooOoqEHysuhiBOmjP3gDv5iXzDkUICCa4jGKIATaYtyJlzUCHPjFFuiTQHdfzADdRZdI1JiYPvugtdyJbAiXnFGkIjxxKlRaCI1DCL1zbgoEkrYmdzssUKmFmgxd+NQkqc4L9aolVkLvGmetw/y/9Y1hKhap3OWBbHzyYybX7uoaX1OExMcyCK8of+OEipk3/28r/0aym9WLpuiXIq3L7DFB+o5V3qK8JtXsPqqS8WDZ4JoHenw6zVPwVSvxXHqqLR9ThJ0amnUptiNmweairnpOoltCL1GCB+IVdAmTfzTDgGgaOBNROqmqxtW0RYUyPJQNaS0RumOQkkikKHGnp3brNIL71HEJwGOSXrZe1ZEtBt6APGzjqROCPYjSJNK2yACzkFyCBcrbkc9FyV4AbWXFexOw5l/0WeFZGdPtEUNcS2JceYx3IO9A==',
        'page_uid': 'ig0m5dp0J1sss4rZRedssssst3N-311207',
    }

    headers = {
        'authority': 'search.naver.com',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'cache-control': 'max-age=0',
        # 'cookie': 'NNB=EOQEJJQUFMYWK; nx_ssl=2; ASID=3d6968c30000018b481054af00008e99; _ga=GA1.2.1867842095.1698124262; _fbp=fb.1.1698124262166.1924394679; _tt_enable_cookie=1; _ttp=BEmRKkFpqC2qDqQF7ScXQ2-Dz-t; _ga_4BKHBFKFK0=GS1.1.1698124262.1.1.1698124263.59.0.0; nid_inf=-1418340302; NID_AUT=3SLnksOBVKphncVoHj0APDbP5U/ED2YpQYleT1yokMIn+HQL9NuWvfzlBs8Vg6FE; NID_JKL=UWjHrUZIdvQ8BXw7yulMx+BHWLqeT5p9oA1MkpzzkAw=; _naver_usersession_=UQOaZKGIfz29lcXjqQmiSA==; NID_SES=AAABpCj5RNSPm+yCjzkae5bDyEbeQ5NrJ1dEqogMrSDUDKoyonJClHDOWMCc3KZL/ANrcNVFE+tRQm/JmswWWNzMPEcBZdWtRzlyrQK8bNWTtco1LekPfTTR6VBlFJxBw/4596ANGrhHafooOoqEHysuhiBOmjP3gDv5iXzDkUICCa4jGKIATaYtyJlzUCHPjFFuiTQHdfzADdRZdI1JiYPvugtdyJbAiXnFGkIjxxKlRaCI1DCL1zbgoEkrYmdzssUKmFmgxd+NQkqc4L9aolVkLvGmetw/y/9Y1hKhap3OWBbHzyYybX7uoaX1OExMcyCK8of+OEipk3/28r/0aym9WLpuiXIq3L7DFB+o5V3qK8JtXsPqqS8WDZ4JoHenw6zVPwVSvxXHqqLR9ThJ0amnUptiNmweairnpOoltCL1GCB+IVdAmTfzTDgGgaOBNROqmqxtW0RYUyPJQNaS0RumOQkkikKHGnp3brNIL71HEJwGOSXrZe1ZEtBt6APGzjqROCPYjSJNK2yACzkFyCBcrbkc9FyV4AbWXFexOw5l/0WeFZGdPtEUNcS2JceYx3IO9A==; page_uid=ig0m5dp0J1sss4rZRedssssst3N-311207',
        'referer': 'https://search.naver.com/search.naver?where=news&sm=tab_pge&query=%EC%82%BC%EC%84%B1%EC%A0%84%EC%9E%90&sort=1&photo=0&field=0&pd=12&ds=2023.10.28.17.56&de=2023.10.28.23.56&mynews=0&office_type=0&office_section_code=0&news_office_checked=&office_category=0&service_area=0&nso=so:dd,p:all,a:all&start=11',
        'sec-ch-ua': '"Chromium";v="118", "Google Chrome";v="118", "Not=A?Brand";v="99"',
        'sec-ch-ua-arch': '"x86"',
        'sec-ch-ua-bitness': '"64"',
        'sec-ch-ua-full-version-list': '"Chromium";v="118.0.5993.117", "Google Chrome";v="118.0.5993.117", "Not=A?Brand";v="99.0.0.0"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-model': '""',
        'sec-ch-ua-platform': '"Windows"',
        'sec-ch-ua-platform-version': '"10.0.0"',
        'sec-ch-ua-wow64': '?0',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36',
    }

    params = {
        'where': 'news',
        'query': inputData['keyword'],
        'sm': 'tab_opt',
        'sort': '1',
        'photo': '0',
        'field': '0',
        'pd': '6',
        'ds': '',
        'de': '',
        'docid': '',
        'related': '0',
        'mynews': '0',
        'office_type': '0',
        'office_section_code': '0',
        'news_office_checked': '',
        'nso': 'so:dd,p:2w',
        'is_sug_officeid': '0',
        'office_category': '0',
        'service_area': '0',
    }
    response = requests.get('https://search.naver.com/search.naver', params=params, cookies=cookies,
                            headers=headers)
    soup=BeautifulSoup(response.text)
    try:
        articleGroup=soup.find("ul",attrs={'class':'list_news'})
        articles = articleGroup.find_all('li', attrs={'class': 'bx'})
    except:
        print("기사없음")
        articles=[]

    articles=articles[:5]
    resultList=[]
    for article in articles:

        title=article.find('a',attrs={'class':'news_tit'}).get_text()
        print("title:",title,"/ title_TYPE:",type(title),len(title))
        url=article.find('a',attrs={'class':'news_tit'})['href']
        print("url:",url,"/ url_TYPE:",type(url),len(url))

        try:
            regiDate=article.find_all('span',attrs={'class':'info'})[-1].get_text().strip()
            regiDate=GetDate(regiDate)
        except:
            regiDate=""
        print("regiDate:",regiDate)

        data={'title':title,'url':url,'regiDate':regiDate}
        resultList.append(data)

    # 데이터를 'regiDate' 키값을 기준으로 최신 순으로 정렬
    resultList = sorted(resultList, key=lambda x: x["regiDate"], reverse=True)

    inputData.update({'news':resultList})

    return inputData



def GetGoogleSpreadSheet():
    scope = 'https://spreadsheets.google.com/feeds'
    json = 'credential.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json, scope)
    gc = gspread.authorize(credentials)
    sheet_url = 'https://docs.google.com/spreadsheets/d/1UTuXepWAxTeeZUMKGDI8Gd8U1cXYyEQKLLXUUumZma4/edit#gid=0'
    doc = gc.open_by_url(sheet_url)
    worksheet = doc.worksheet('시트1')
    print(worksheet)
    #=================특정행의 정보 가져오기
    # cell_data = worksheet.acell('A1').value
    #=================전체정보가져오기
    all_data=worksheet.get_all_records()
    #==================맨 밑행에 데이타 넣기
    # new_row = ['John', 30, 'Teacher']
    # worksheet.append_row(new_row)
    pprint.pprint(all_data)
    return all_data


def DoRun():
    #=========구글시트정보가져오기
    while True:
        try:
            spreadDatas=GetGoogleSpreadSheet()
            break
        except:
            print("구글스프레드에러")
            time.sleep(60)

    for spreadData in spreadDatas:
        pprint.pprint(spreadData)
        newsList = []
        pubmedList = []
        try:
            naverName1={'name':spreadData['교수님1(네이버)'],'keyword':spreadData['검색어1(네이버)'],}
        except:
            naverName1=""
        print("naverName1:",naverName1)
        if len(naverName1['name'])>=1:
            newsList.append(naverName1)
        try:
            naverName2={'name':spreadData['교수님2(네이버)'],'keyword':spreadData['검색어2(네이버)'],}
        except:
            naverName2=""
        print("naverName2:",naverName2)
        if len(naverName2['name'])>=1:
            newsList.append(naverName2)
        try:
            naverName3={'name':spreadData['교수님3(네이버)'],'keyword':spreadData['검색어3(네이버)'],}
        except:
            naverName3=""
        print("naverName3:",naverName3)
        if len(naverName3['name'])>=1:
            newsList.append(naverName3)
        try:
            naverName4={'name':spreadData['교수님4(네이버)'],'keyword':spreadData['검색어4(네이버)'],}
        except:
            naverName4=""
        print("naverName4:",naverName4)
        if len(naverName4['name'])>=1:
            newsList.append(naverName4)
        try:
            naverName5={'name':spreadData['교수님5(네이버)'],'keyword':spreadData['검색어5(네이버)'],}
        except:
            naverName5=""
        print("naverName5:",naverName5)
        if len(naverName5['name'])>=1:
            newsList.append(naverName5)


        try:
            pubmedName1={'name':spreadData['교수님1(PUBMED)'],'keyword':spreadData['검색어1(PUBMED)'],}
        except:
            pubmedName1=""
        print("pubmedName1:",pubmedName1)
        if len(pubmedName1['name'])>=1:
            pubmedList.append(pubmedName1)
        try:
            pubmedName2={'name':spreadData['교수님2(PUBMED)'],'keyword':spreadData['검색어2(PUBMED)'],}
        except:
            pubmedName2=""
        print("pubmedName2:",pubmedName2)
        if len(pubmedName2['name'])>=1:
            pubmedList.append(pubmedName2)
        try:
            pubmedName3={'name':spreadData['교수님3(PUBMED)'],'keyword':spreadData['검색어3(PUBMED)'],}
        except:
            pubmedName3=""
        print("pubmedName3:",pubmedName3)
        if len(pubmedName3['name'])>=1:
            pubmedList.append(pubmedName3)
        try:
            pubmedName4={'name':spreadData['교수님4(PUBMED)'],'keyword':spreadData['검색어4(PUBMED)'],}
        except:
            pubmedName4=""
        print("pubmedName4:",pubmedName4)
        if len(pubmedName4['name'])>=1:
            pubmedList.append(pubmedName4)
        try:
            pubmedName5={'name':spreadData['교수님5(PUBMED)'],'keyword':spreadData['검색어5(PUBMED)'],}
        except:
            pubmedName5=""
        print("pubmedName5:",pubmedName5)
        if len(pubmedName5['name'])>=1:
            pubmedList.append(pubmedName5)

        totalNews=[]
        #========뉴스랑 펍메드가져오기
        for index,inputNews in enumerate(newsList):
            newsList=GetNews(inputNews)
            text="{}번째 확인중...".format(index+1)
            print(text)
            totalNews.append(newsList)
            time.sleep(random.randint(10,20)*0.1)
        totalArticles=[]
        for index,inputPubmed in enumerate(pubmedList):
            articleList=GetArticlesPubMed(inputPubmed)
            text = "{}번째 확인중...".format(index+1)
            print(text)
            totalArticles.append(articleList)
            time.sleep(random.randint(10, 20) * 0.1)
        with open('totalNews.json', 'w',encoding='utf-8-sig') as f:
            json.dump(totalNews, f, indent=2,ensure_ascii=False)
        with open('totalArticles.json', 'w',encoding='utf-8-sig') as f:
            json.dump(totalArticles, f, indent=2,ensure_ascii=False)

        #==========메일보내기
        with open ('totalNews.json', "r",encoding='utf-8-sig') as f:
            totalNews = json.load(f)
        with open ('totalArticles.json', "r",encoding='utf-8-sig') as f:
            totalArticles = json.load(f)
        send_email(totalNews,totalArticles,spreadData)

def NowYoYil():
    # 현재 날짜와 시간을 가져옵니다.
    now = datetime.datetime.now()
    # 요일을 한글로 변환하는 딕셔너리를 만듭니다.
    weekday_dict = {
        0: '월',
        1: '화',
        2: '수',
        3: '목',
        4: '금',
        5: '토',
        6: '일'
    }
    # 현재 날짜의 요일을 숫자로 가져옵니다. (월요일은 0, 일요일은 6)
    today_weekday_number = now.weekday()
    # 숫자를 한글 요일로 변환합니다.
    today_weekday_korean = weekday_dict[today_weekday_number]
    return today_weekday_korean


while True:
    try:
        keywordList = GetGoogleSpreadSheet()
        break
    except:
        print("구글스프레드에러")
        time.sleep(60)

# 함수를 예약합니다. (예: 매일 오후 3시)
while True:
    timeNowString=datetime.datetime.now().strftime("%H%M%S")
    # timeTarget=datetime.datetime.now().strftime("%Y%m%d_{}{}{}".format(keywordList[0]['발송시간'],'00','00'))
    timeNow=datetime.datetime.now()
    regex=re.compile('\d+')
    sendTimeNumber=regex.findall(keywordList[0]['송신시간'])[0]
    sendTimeString= re.sub(r"[^가-힣]", "", keywordList[0]['송신시간'])  # 자모가 아닌 한글만 남기기(공백 제거)
    resultNowYoYil=NowYoYil()
    timeTarget=dt = datetime.datetime(timeNow.year,timeNow.month,timeNow.day,int(sendTimeNumber), 0, 0).strftime("%H%M%S")
    text="시간:{}/{},요일:{}/{}, 송신예약:{}".format(timeNowString,timeTarget,sendTimeString,resultNowYoYil,sendTimeString+str(sendTimeNumber))
    print(text)

    # if timeNowString==timeTarget and sendTimeString==resultNowYoYil:
    if timeNowString==timeTarget:
    # if True:
        DoRun()
    # break
    time.sleep(1)
