<p>
This script demonstrates how to retrieve the Office 365 Service Health Data using the <strong><u>Office 365 Management API</u></strong>, and send the email report using <strong><u>Microsoft Graph API</u></strong>.</p>
<h3>

</h3>
<h3>
What is covered by this post?</h3>
<ul>
<li>App Registration in Azure AD</li>
<li>Configuring the Script</li>
<li>Running the Script and Generating Outputs / Reports</li>
</ul>
<p>

</p>
<h3>
What is NOT covered by this post?</h3>
<p>
This post does not cover the “<em>How-To”</em> of the said APIs, because they can already be found by following these links: </p>
<ul>
<li><strong>Office 365 Management APIs</strong> - <a title="https://docs.microsoft.com/en-us/office/office-365-management-api/" href="https://docs.microsoft.com/en-us/office/office-365-management-api/">https://docs.microsoft.com/en-us/office/office-365-management-api/</a></li>
<li><strong>Microsoft Graph API</strong> - <a title="https://developer.microsoft.com/en-us/graph/docs/concepts/overview" href="https://developer.microsoft.com/en-us/graph/docs/concepts/overview">https://developer.microsoft.com/en-us/graph/docs/concepts/overview</a></li>
</ul>
<p>

</p>
<h3>
Requirements</h3>
<ul>
<li>Application Registration in Azure AD (Application ID + Key + Permissions)</li>
<li>Exchange Online Mailbox (User or Shared Mailbox, for sending reports)</li>
</ul>
<p>

</p>
<h3>
Script Download</h3>
<p>
<a title="https://github.com/junecastillote/Get-O365HealthReport/archive/master.zip" href="https://github.com/junecastillote/Get-O365HealthReport/archive/master.zip">https://github.com/junecastillote/Get-O365HealthReport/archive/master.zip</a></p>
<p>

</p>
<h3>
App Registration</h3>
<p>
Note: Your account must be a Global Admin</p>
<ul>
<li>Go to <strong>Azure Active Directory</strong> &gt; <strong>App Registrations</strong></li>
</ul>
<blockquote>
<p>
&nbsp;<a href="https://lh3.googleusercontent.com/-6P8YGk-OtIk/W-KowVXXasI/AAAAAAAADlY/FZpMMi4I4BsIXTAYr5T5ZhuMs6aHYmGBACHMYCw/s1600-h/mRemoteNG_2018-11-07_13-43-15%255B6%255D" target="_blank"><img width="485" height="104" title="" style="border: 0px currentcolor; border-image: none; display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-F9MarjzInQY/W-Koxa62UHI/AAAAAAAADlc/jgJ5siSfWwcmxhyFyhljRKaEDvAD8981wCHMYCw/mRemoteNG_2018-11-07_13-43-15_thumb%255B4%255D?imgmax=800" border="0"></a></p>
</blockquote>
<ul>
<li>Click <strong>New Application Registration</strong></li>
<li>Fill out the <strong>Name</strong>, <strong>Application Type</strong> and <strong>Sign-on URL</strong> as shown below, then click <strong>Create</strong></li>
</ul>
<blockquote>
<p>
&nbsp;<a href="https://lh3.googleusercontent.com/-N9Xcy8QFwDM/W-Koyk7VqQI/AAAAAAAADlg/IENcgBSqv3MTjgkAx0KvAbV8FTu69uZagCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-47-33%255B5%255D" target="_blank"><img width="321" height="237" title="" style="border: 0px currentcolor; border-image: none; display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-BICaLFkiaU0/W-KozpYXCHI/AAAAAAAADlk/3XJHBaPuEogARCBh36TshyS9YSb2K2MBQCHMYCw/mRemoteNG_2018-11-07_13-47-33_thumb%255B3%255D?imgmax=800" border="0"></a></p>
</blockquote>
<ul>
<li>Once the App is registered, copy the<strong> Application ID</strong> for later use.</li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-jr2V-8XVanY/W-Ko0s0LcwI/AAAAAAAADlo/CXjmOPA9U7sw-6zS4fzSYBuDmKAEa6cXgCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-48-51%255B9%255D" target="_blank"><img width="573" height="233" title="" style="border: 0px currentcolor; border-image: none; display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-SNis-A0I3Eo/W-Ko12pWdaI/AAAAAAAADls/CY4rtOx_gv8GFIlmQfoahGXiJJroyYO_gCHMYCw/mRemoteNG_2018-11-07_13-48-51_thumb%255B7%255D?imgmax=800" border="0"></a>&nbsp;</p>
</blockquote>
<ul>
<li>Click <strong>Settings</strong> &gt; <strong>Keys</strong></li>
</ul>
<blockquote>
<p>
<strong><a href="https://lh3.googleusercontent.com/-LLp1b4OovnE/W-Ko2wvAAUI/AAAAAAAADlw/wkN4oA_d6VUqOMq5Alf6oI4v_4v1gZiwwCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-04-14%255B4%255D" target="_blank"><img width="300" height="129" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-b7vwLC4wUUA/W-Ko3q1uMpI/AAAAAAAADl0/jXeeM6Ad6wAox1D5hAw-iuV_Wf5RfLOGQCHMYCw/mRemoteNG_2018-11-07_14-04-14_thumb%255B2%255D?imgmax=800" border="0"></a></strong></p>
</blockquote>
<ul>
<li>Type in the <strong>Description</strong> and select the <strong>expiration</strong> for your key, then click <strong>Save</strong></li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-3ZdNu_agy6o/W-Ko4zzVzGI/AAAAAAAADl4/hYexRyUA5BQ_gldkGjSXa1tYzucidpXMwCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-54-50%255B6%255D" target="_blank"><img width="453" height="211" title="" style="border: 0px currentcolor; border-image: none; display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-Q94STD9m3nE/W-Ko6QRILMI/AAAAAAAADl8/P7oBUnvK3P8shD736dNGbh3KqfNXJL5EACHMYCw/mRemoteNG_2018-11-07_13-54-50_thumb%255B4%255D?imgmax=800" border="0"></a>&nbsp;</p>
</blockquote>
<ul>
<li>After clicking Save, the Key will be generated. You must copy this key value because it will not be shown again. </li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-2TF3u51UG08/W-Ko7f6wa5I/AAAAAAAADmA/exhLDdeWchg4zyKU-lRFX9WF8oYy51AAQCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-59-21%255B8%255D" target="_blank"><img width="637" height="79" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-FJbwsNjeQuY/W-Ko8U4Sz9I/AAAAAAAADmE/CsGTasPe_lAXuK8mFD7wTksbMFfbgUSKgCHMYCw/mRemoteNG_2018-11-07_13-59-21_thumb%255B4%255D?imgmax=800" border="0"></a></p>
</blockquote>
<ul>
<li>Go to Required Permissions</li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-suYwUJOCTds/W-Ko9RPFnMI/AAAAAAAADmI/3fcC5G1p1hM5FGl5BrA-IYFyag-P1ndpACHMYCw/s1600-h/mRemoteNG_2018-11-07_14-04-56%255B5%255D" target="_blank"><img width="297" height="124" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-Tujh2Es549U/W-Ko-m7UtJI/AAAAAAAADmM/yPtDtssKNfYUW9-vEuTdIHMOhKYo_WPFwCHMYCw/mRemoteNG_2018-11-07_14-04-56_thumb%255B3%255D?imgmax=800" border="0"></a>&nbsp;</p>
</blockquote>
<ul>
<li>By default, the <strong>Windows Azure Active Directory</strong> API is already added. <strong>This is not needed and should be deleted.</strong></li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-IaL07SveDTI/W-Ko_j_EVeI/AAAAAAAADmQ/LzoehfY30vIhMGUiR-2uClSxnUvkYMqoQCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-06-51%255B4%255D" target="_blank"><img width="565" height="99" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-L73f0k3dmfY/W-KpBEAxehI/AAAAAAAADmU/D2khLYYABm08pvDKPQdXUt2lFFB_cAdlgCHMYCw/mRemoteNG_2018-11-07_14-06-51_thumb%255B2%255D?imgmax=800" border="0"></a>&nbsp;</p>
<p>
<a href="https://lh3.googleusercontent.com/-4deAzz1b3E8/W-KpCFsJ-6I/AAAAAAAADmY/5vQe4AXMfucC0lu7blDB06Jqyq8SF76OgCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-07-25%255B4%255D" target="_blank"><img width="232" height="91" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-E2lEMWsJPx4/W-KpCzkisBI/AAAAAAAADmc/QfNaDppcJtE7ciM5to4aePLQis1K4EJSgCHMYCw/mRemoteNG_2018-11-07_14-07-25_thumb%255B2%255D?imgmax=800" border="0"></a></p>
</blockquote>
<ul>
<li>Add the following APIs and Permissions</li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-jEfK20mK9Fs/W-KpENu2a3I/AAAAAAAADmg/0bFvJKyVhaAzoaZZVsEBO6HklaDRoOPiQCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-08-54%255B4%255D" target="_blank"><img width="238" height="96" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-xnUBKGA-8Ck/W-KpFGyoO0I/AAAAAAAADmo/s3BTqkIyAlUQ1ov6kH7ryXSr_fmxNh6BwCHMYCw/mRemoteNG_2018-11-07_14-08-54_thumb%255B2%255D?imgmax=800" border="0"></a></p>
</blockquote>
<p>
<strong>Office 365 Management APIs</strong></p>
<p>
<a href="https://lh3.googleusercontent.com/-n5H5QTmWlds/W-KpGnN799I/AAAAAAAADms/fwfX0KlJG3AyRplxpySDJ8Xmhxfdvd45QCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-09-55%255B4%255D"><img width="831" height="356" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-07hrZmBhPnU/W-KpHtg_85I/AAAAAAAADmw/kYWLMzC0Gnc9ahi1erhTjFRRuFn0sgI8QCHMYCw/mRemoteNG_2018-11-07_14-09-55_thumb%255B2%255D?imgmax=800" border="0"></a>&nbsp;</p>
<p>
<strong>Microsoft Graph API</strong></p>
<p>
<a href="https://lh3.googleusercontent.com/-KTjLlv7ZjgM/W-KpIk4y1EI/AAAAAAAADm0/BYm2znds_fQHoIG9gvF11d9cD_I8UHUOQCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-11-59%255B6%255D"><img width="832" height="218" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-4a0lktMy5YM/W-KpJ6EJo3I/AAAAAAAADm4/E5M-iNg5FWYKzH4dTbleZPFeHBaPUpgvwCHMYCw/mRemoteNG_2018-11-07_14-11-59_thumb%255B4%255D?imgmax=800" border="0"></a></p>
<ul>
<li>Once Required Permissions are added, click <strong>Grant Permissions</strong></li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-WEvxpxD8AdI/W-KpLJz3QtI/AAAAAAAADm8/ITDcGnNUhmQPA3ap8v7ISTpkRwz_gUopgCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-14-14%255B4%255D" target="_blank"><img width="468" height="209" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-aUfv4-e2_88/W-KpMN_Lm6I/AAAAAAAADnA/mNZgtPSDayIZyo1JGRW-_5j4IcXmj611gCHMYCw/mRemoteNG_2018-11-07_14-14-14_thumb%255B2%255D?imgmax=800" border="0"></a>&nbsp;</p>
</blockquote>
<ul>
<li>Click <strong>Yes</strong></li>
</ul>
<blockquote>
<p>
<a href="https://lh3.googleusercontent.com/-_c1lFxX-dy0/W-KpNVJCaoI/AAAAAAAADnE/zT1IfBt55_ACbV9X-gEIkxadzqSiP_eLACHMYCw/s1600-h/mRemoteNG_2018-11-07_14-16-59%255B4%255D" target="_blank"><img width="540" height="165" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-ZjPNBez7MIg/W-KpOct3YdI/AAAAAAAADnM/XmvEcsd4JcYTA0qcfibAiPWYUlxtT_fbgCHMYCw/mRemoteNG_2018-11-07_14-16-59_thumb%255B2%255D?imgmax=800" border="0"></a>&nbsp;</p>
</blockquote>
<h3>
Script Configuration</h3>
<p>
Open the config.xml file and edit the values as necessary like the example below:</p>
<p>
<a href="https://lh3.googleusercontent.com/-c0lwSC3sQjM/W-KpP_7QLyI/AAAAAAAADnQ/LBfkhZKiyJQNs7p2ZhW9jWs-SAJA6V6WQCHMYCw/s1600-h/mRemoteNG_2018-11-07_16-06-47%255B3%255D" target="_blank"><img width="790" height="212" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-05-1L1WH3yo/W-KpQzxSGGI/AAAAAAAADnU/QYo9BoPssXUQdqnra6UG4CIdNnFcv0fGgCHMYCw/mRemoteNG_2018-11-07_16-06-47_thumb%255B1%255D?imgmax=800" border="0"></a></p>
<p>
<strong>sendEmail</strong> – set this to TRUE or FALSE depending on whether you want the report sent thru email.</p>
<p>
<strong>testMode</strong> – set this to TRUE or FALSE depending on whether you want to run in test mode or not. Test Mode will treat ALL items retrieved from the service health dashboard as NEW or UPDATE. When you’re ready to put this script in production, set this to FALSE</p>
<p>
<strong>clientID</strong> – this is the Application ID you copied from the App Registration in Azure AD</p>
<p>
<strong>clientSecret</strong> – this is the Key you copied from the App Registration in Azure AD</p>
<p>
<strong>tenantDomain</strong> – this is your Office 365 Tenant Domain</p>
<p>
<strong>toAddress</strong> – your intended recipients of the report, separate multiple recipients with a comma with no spaces.</p>
<p>
<strong>fromAddress</strong> – the primary smtp address of the Shared Mailbox or User Mailbox you want to use for sending the email report.</p>
<p>
<strong>mailSubject</strong> – you preferred email subject and title of the html report</p>
<p>

</p>
<h3>
Running the Script</h3>
<p>
<strong>IMPORTANT</strong>: In the first run, whether in Test Mode or not, will only generate the data that will be needed for future run comparisons.</p>
<p>
In this example, the script is in run Test Mode.</p>
<p>
<a href="https://lh3.googleusercontent.com/-GPWkIajQZDw/W-KpSLQ2FCI/AAAAAAAADnY/w476Wjg81241nEnS_8pS6-7NEs-jMG2vgCHMYCw/s1600-h/mRemoteNG_2018-11-07_16-24-26%255B4%255D" target="_blank"><img width="542" height="201" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-a_VpPCmCW2k/W-KpTYb-QII/AAAAAAAADnc/wKvTtweTYUkZ0ia--kDvN5x0Sv1yP6TPwCHMYCw/mRemoteNG_2018-11-07_16-24-26_thumb%255B2%255D?imgmax=800" border="0"></a></p>
<h3>
Sample Output</h3>
<h4>
Email</h4>
<p>
<a href="https://lh3.googleusercontent.com/-lhlI0oR1hX0/W-KpUjIqNsI/AAAAAAAADnk/NcDiTQYMH8kenSIAtp2GeBuy9GjFXq8WACHMYCw/s1600-h/mRemoteNG_2018-11-07_16-40-40%255B8%255D" target="_blank"><img width="970" height="617" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-6-Tgek7_p7I/W-KpVlIYTEI/AAAAAAAADno/HeNxxsKkHBYVSZE1KjoV0vLghAY8Gxe4QCHMYCw/mRemoteNG_2018-11-07_16-40-40_thumb%255B6%255D?imgmax=800" border="0"></a></p>
<h4>
HTML</h4>
<p>
<a href="https://lh3.googleusercontent.com/-QhTrSYxScTw/W-KpW7BzmvI/AAAAAAAADns/gXvS3ftrB9cT7ZTC9TNgOU4xlnI6Z9EwACHMYCw/s1600-h/mRemoteNG_2018-11-07_16-42-12%255B3%255D"><img width="422" height="180" title="mRemoteNG_2018-11-07_16-42-12" style="display: inline; background-image: none;" alt="mRemoteNG_2018-11-07_16-42-12" src="https://lh3.googleusercontent.com/-5lnD5qEUEBE/W-KpYGawzkI/AAAAAAAADnw/jVRlh7oknMkOoOAo4TcBm5_iZXhyVJMMACHMYCw/mRemoteNG_2018-11-07_16-42-12_thumb%255B1%255D?imgmax=800" border="0"></a></p>
<p>
<a href="https://lh3.googleusercontent.com/-iJB_hO5SCwE/W-KpZTfKtrI/AAAAAAAADn0/jigBS5MSImUBU6HDr5PN91i_Jm5LOa3-gCHMYCw/s1600-h/mRemoteNG_2018-11-07_16-44-53%255B2%255D"><img width="670" height="582" title="mRemoteNG_2018-11-07_16-44-53" style="display: inline; background-image: none;" alt="mRemoteNG_2018-11-07_16-44-53" src="https://lh3.googleusercontent.com/-sH_lBTOGtdg/W-KpakkAtgI/AAAAAAAADn8/gsb0Qfo9B6EfQhhSu8TXHpO-rrKKBHNxwCHMYCw/mRemoteNG_2018-11-07_16-44-53_thumb?imgmax=800" border="0"></a></p>
<p>
This script is functional, but I’m sure there can be many improvements that can be done. So please feel free to comment or modify, just please don’t forget to credit the original source.</p>
