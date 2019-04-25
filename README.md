This script demonstrates how to retrieve the Office 365 Service Health Data using the <strong><u>Office 365 Management API</u></strong>, and send the email report using <strong><u>Microsoft Graph API</u></strong>.
<strong><u>
</u></strong> <strong><u>The logic flow is simple:</u></strong>
<ol>
<li>Retrieve Office 365 Service Health Record (this is the only action done during the first run, saved to new.csv)</li>
<li>Read Old Records from the file (old.csv)</li>
<li>Compare retrieved records with old records (new.csv VS old.csv)</li>
<li>Report if there are new or updated records (updated.csv)</li>
</ol>
You may want to have this running as a scheduled task at an interval you prefer.
<h3>
What is covered by this post?</h3>
<ul>
<li>App Registration in Azure AD</li>
<li>Configuring the Script</li>
<li>Running the Script and Generating Outputs / Reports</li>
</ul>
<h3>
What is NOT covered by this post?</h3>
This post does not cover the “<em>How-To”</em> of the said APIs, because they can already be found by following these links:
<ul>
<li><strong>Office 365 Management APIs</strong> - <a href="https://docs.microsoft.com/en-us/office/office-365-management-api/" title="https://docs.microsoft.com/en-us/office/office-365-management-api/">https://docs.microsoft.com/en-us/office/office-365-management-api/</a></li>
<li><strong>Microsoft Graph API</strong> - <a href="https://developer.microsoft.com/en-us/graph/docs/concepts/overview" title="https://developer.microsoft.com/en-us/graph/docs/concepts/overview">https://developer.microsoft.com/en-us/graph/docs/concepts/overview</a></li>
</ul>
<h3>
Requirements</h3>
<ul>
<li>Application Registration in Azure AD (Application ID + Key + Permissions)</li>
<li>Exchange Online Mailbox (User or Shared Mailbox, for sending reports)Script Download /&nbsp; Change Logs</li>
</ul>
v1.5 (latest) -&nbsp;<a href="https://github.com/junecastillote/Get-O365HealthReport">https://github.com/junecastillote/Get-O365HealthReport</a>
<ul>
<li>code cleanup</li>
<li>added code to cater to the new App Registration Keys special characters</li>
</ul>
v1.4
<ul>
<li>code cleanup</li>
<li>fixed JSON conversion for email report</li>
</ul>
v1.3
<ul>
<li>added "exclusion" feature. (requested from&nbsp;this <a href="https://github.com/junecastillote/Get-O365HealthReport/issues/1">issue</a>)</li>
<li>the&nbsp;<b>exclusions.csv</b> file inside the <b>\resource</b> folder can now be used to exclude workloads from the report.</li>
</ul>
v1.2
<ul>
<li>Modified to also check the changes in "Status" to trigger an update alert. (eg. Service Degradation to Service Restored). This is because I observed that some events' Last Updated Time does not change but the Status change which is not getting captured by the previous script.</li>
</ul>
v1.1
<ul>
<li>Added “organizationName” field in config.xml</li>
<li>Removed “mailSubject” field from config.xml</li>
<li>Send one email per event (alerts are no longer consolidated in one single email)</li>
</ul>
v1.0

<ul>
<li>Initial build</li>
</ul>


<h3>
App Registration</h3>
Note: Your account must be a Global Admin
<ul>
<li>Go to <strong>Azure Active Directory</strong> &gt; <strong>App Registrations</strong></li>
</ul>
<blockquote>
&nbsp;<a href="https://lh3.googleusercontent.com/-6P8YGk-OtIk/W-KowVXXasI/AAAAAAAADlY/FZpMMi4I4BsIXTAYr5T5ZhuMs6aHYmGBACHMYCw/s1600-h/mRemoteNG_2018-11-07_13-43-15%255B6%255D" target="_blank"><img alt="" border="0" height="104" src="https://lh3.googleusercontent.com/-F9MarjzInQY/W-Koxa62UHI/AAAAAAAADlc/jgJ5siSfWwcmxhyFyhljRKaEDvAD8981wCHMYCw/mRemoteNG_2018-11-07_13-43-15_thumb%255B4%255D?imgmax=800" style="background-image: none; border: 0px currentcolor; display: inline;" title="" width="485" /></a></blockquote>
<ul>
<li>Click <strong>New Application Registration</strong></li>
<li>Fill out the <strong>Name</strong>, <strong>Application Type</strong> and <strong>Sign-on URL</strong> as shown below, then click <strong>Create</strong></li>
</ul>
<blockquote>
&nbsp;<a href="https://lh3.googleusercontent.com/-N9Xcy8QFwDM/W-Koyk7VqQI/AAAAAAAADlg/IENcgBSqv3MTjgkAx0KvAbV8FTu69uZagCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-47-33%255B5%255D" target="_blank"><img alt="" border="0" height="237" src="https://lh3.googleusercontent.com/-BICaLFkiaU0/W-KozpYXCHI/AAAAAAAADlk/3XJHBaPuEogARCBh36TshyS9YSb2K2MBQCHMYCw/mRemoteNG_2018-11-07_13-47-33_thumb%255B3%255D?imgmax=800" style="background-image: none; border: 0px currentcolor; display: inline;" title="" width="321" /></a></blockquote>
<ul>
<li>Once the App is registered, copy the<strong> Application ID</strong> for later use.</li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-jr2V-8XVanY/W-Ko0s0LcwI/AAAAAAAADlo/CXjmOPA9U7sw-6zS4fzSYBuDmKAEa6cXgCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-48-51%255B9%255D" target="_blank"><img alt="" border="0" height="233" src="https://lh3.googleusercontent.com/-SNis-A0I3Eo/W-Ko12pWdaI/AAAAAAAADls/CY4rtOx_gv8GFIlmQfoahGXiJJroyYO_gCHMYCw/mRemoteNG_2018-11-07_13-48-51_thumb%255B7%255D?imgmax=800" style="background-image: none; border: 0px currentcolor; display: inline;" title="" width="573" /></a>&nbsp;</blockquote>
<ul>
<li>Click <strong>Settings</strong> &gt; <strong>Keys</strong></li>
</ul>
<blockquote>
<strong><a href="https://lh3.googleusercontent.com/-LLp1b4OovnE/W-Ko2wvAAUI/AAAAAAAADlw/wkN4oA_d6VUqOMq5Alf6oI4v_4v1gZiwwCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-04-14%255B4%255D" target="_blank"><img alt="" border="0" height="129" src="https://lh3.googleusercontent.com/-b7vwLC4wUUA/W-Ko3q1uMpI/AAAAAAAADl0/jXeeM6Ad6wAox1D5hAw-iuV_Wf5RfLOGQCHMYCw/mRemoteNG_2018-11-07_14-04-14_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="300" /></a></strong></blockquote>
<ul>
<li>Type in the <strong>Description</strong> and select the <strong>expiration</strong> for your key, then click <strong>Save</strong></li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-3ZdNu_agy6o/W-Ko4zzVzGI/AAAAAAAADl4/hYexRyUA5BQ_gldkGjSXa1tYzucidpXMwCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-54-50%255B6%255D" target="_blank"><img alt="" border="0" height="211" src="https://lh3.googleusercontent.com/-Q94STD9m3nE/W-Ko6QRILMI/AAAAAAAADl8/P7oBUnvK3P8shD736dNGbh3KqfNXJL5EACHMYCw/mRemoteNG_2018-11-07_13-54-50_thumb%255B4%255D?imgmax=800" style="background-image: none; border: 0px currentcolor; display: inline;" title="" width="453" /></a>&nbsp;</blockquote>
<ul>
<li>After clicking Save, the Key will be generated. You must copy this key value because it will not be shown again.</li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-2TF3u51UG08/W-Ko7f6wa5I/AAAAAAAADmA/exhLDdeWchg4zyKU-lRFX9WF8oYy51AAQCHMYCw/s1600-h/mRemoteNG_2018-11-07_13-59-21%255B8%255D" target="_blank"><img alt="" border="0" height="79" src="https://lh3.googleusercontent.com/-FJbwsNjeQuY/W-Ko8U4Sz9I/AAAAAAAADmE/CsGTasPe_lAXuK8mFD7wTksbMFfbgUSKgCHMYCw/mRemoteNG_2018-11-07_13-59-21_thumb%255B4%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="637" /></a></blockquote>
<ul>
<li>Go to Required Permissions</li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-suYwUJOCTds/W-Ko9RPFnMI/AAAAAAAADmI/3fcC5G1p1hM5FGl5BrA-IYFyag-P1ndpACHMYCw/s1600-h/mRemoteNG_2018-11-07_14-04-56%255B5%255D" target="_blank"><img alt="" border="0" height="124" src="https://lh3.googleusercontent.com/-Tujh2Es549U/W-Ko-m7UtJI/AAAAAAAADmM/yPtDtssKNfYUW9-vEuTdIHMOhKYo_WPFwCHMYCw/mRemoteNG_2018-11-07_14-04-56_thumb%255B3%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="297" /></a>&nbsp;</blockquote>
<ul>
<li>By default, the <strong>Windows Azure Active Directory</strong> API is already added. <strong>This is not needed and should be deleted.</strong></li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-IaL07SveDTI/W-Ko_j_EVeI/AAAAAAAADmQ/LzoehfY30vIhMGUiR-2uClSxnUvkYMqoQCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-06-51%255B4%255D" target="_blank"><img alt="" border="0" height="99" src="https://lh3.googleusercontent.com/-L73f0k3dmfY/W-KpBEAxehI/AAAAAAAADmU/D2khLYYABm08pvDKPQdXUt2lFFB_cAdlgCHMYCw/mRemoteNG_2018-11-07_14-06-51_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="565" /></a>
<a href="https://lh3.googleusercontent.com/-4deAzz1b3E8/W-KpCFsJ-6I/AAAAAAAADmY/5vQe4AXMfucC0lu7blDB06Jqyq8SF76OgCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-07-25%255B4%255D" target="_blank"><img alt="" border="0" height="91" src="https://lh3.googleusercontent.com/-E2lEMWsJPx4/W-KpCzkisBI/AAAAAAAADmc/QfNaDppcJtE7ciM5to4aePLQis1K4EJSgCHMYCw/mRemoteNG_2018-11-07_14-07-25_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="232" /></a></blockquote>
<ul>
<li>Add the following APIs and Permissions</li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-jEfK20mK9Fs/W-KpENu2a3I/AAAAAAAADmg/0bFvJKyVhaAzoaZZVsEBO6HklaDRoOPiQCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-08-54%255B4%255D" target="_blank"><img alt="" border="0" height="96" src="https://lh3.googleusercontent.com/-xnUBKGA-8Ck/W-KpFGyoO0I/AAAAAAAADmo/s3BTqkIyAlUQ1ov6kH7ryXSr_fmxNh6BwCHMYCw/mRemoteNG_2018-11-07_14-08-54_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="238" /></a></blockquote>
<strong>Office 365 Management APIs</strong>
<a href="https://lh3.googleusercontent.com/-n5H5QTmWlds/W-KpGnN799I/AAAAAAAADms/fwfX0KlJG3AyRplxpySDJ8Xmhxfdvd45QCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-09-55%255B4%255D"><img alt="" border="0" height="356" src="https://lh3.googleusercontent.com/-07hrZmBhPnU/W-KpHtg_85I/AAAAAAAADmw/kYWLMzC0Gnc9ahi1erhTjFRRuFn0sgI8QCHMYCw/mRemoteNG_2018-11-07_14-09-55_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="831" /></a>
<strong>Microsoft Graph API</strong>
<a href="https://lh3.googleusercontent.com/-KTjLlv7ZjgM/W-KpIk4y1EI/AAAAAAAADm0/BYm2znds_fQHoIG9gvF11d9cD_I8UHUOQCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-11-59%255B6%255D"><img alt="" border="0" height="218" src="https://lh3.googleusercontent.com/-4a0lktMy5YM/W-KpJ6EJo3I/AAAAAAAADm4/E5M-iNg5FWYKzH4dTbleZPFeHBaPUpgvwCHMYCw/mRemoteNG_2018-11-07_14-11-59_thumb%255B4%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="832" /></a>
<ul>
<li>Once Required Permissions are added, click <strong>Grant Permissions</strong></li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-WEvxpxD8AdI/W-KpLJz3QtI/AAAAAAAADm8/ITDcGnNUhmQPA3ap8v7ISTpkRwz_gUopgCHMYCw/s1600-h/mRemoteNG_2018-11-07_14-14-14%255B4%255D" target="_blank"><img alt="" border="0" height="209" src="https://lh3.googleusercontent.com/-aUfv4-e2_88/W-KpMN_Lm6I/AAAAAAAADnA/mNZgtPSDayIZyo1JGRW-_5j4IcXmj611gCHMYCw/mRemoteNG_2018-11-07_14-14-14_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="468" /></a>&nbsp;</blockquote>
<ul>
<li>Click <strong>Yes</strong></li>
</ul>
<blockquote>
<a href="https://lh3.googleusercontent.com/-_c1lFxX-dy0/W-KpNVJCaoI/AAAAAAAADnE/zT1IfBt55_ACbV9X-gEIkxadzqSiP_eLACHMYCw/s1600-h/mRemoteNG_2018-11-07_14-16-59%255B4%255D" target="_blank"><img alt="" border="0" height="165" src="https://lh3.googleusercontent.com/-ZjPNBez7MIg/W-KpOct3YdI/AAAAAAAADnM/XmvEcsd4JcYTA0qcfibAiPWYUlxtT_fbgCHMYCw/mRemoteNG_2018-11-07_14-16-59_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="540" /></a>&nbsp;</blockquote>
<h3>
Script Configuration</h3>
Open the config.xml file and edit the values as necessary like the example below:
<a href="https://lh3.googleusercontent.com/--A4Ef_1h-hM/W_4f9PN75XI/AAAAAAAAEKU/xHbZvmNfiF0gpaCQfUPYPlpxqz-1zyBfQCHMYCw/s1600-h/mRemoteNG_2018-11-28_12-45-24%255B7%255D" target="_blank"><img alt="" border="0" height="208" src="https://lh3.googleusercontent.com/-jaSnCfiWJ-M/W_4f_TF4hSI/AAAAAAAAEKY/3Kj_fPve6SE8hqCK9saWmhOOtazz3-qAwCHMYCw/mRemoteNG_2018-11-28_12-45-24_thumb%255B4%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="774" /></a>
<strong>
</strong> <strong>sendEmail</strong> – set this to TRUE or FALSE depending on whether you want the report sent thru email.
<strong>testMode</strong> – set this to TRUE or FALSE depending on whether you want to run in test mode or not. Test Mode will treat ALL items retrieved from the service health dashboard as NEW or UPDATE. When you’re ready to put this script in production, set this to FALSE
<strong>clientID</strong> – this is the Application ID you copied from the App Registration in Azure AD
<strong>clientSecret</strong> – this is the Key you copied from the App Registration in Azure AD
<strong>tenantDomain</strong> – this is your Office 365 Tenant Domain
<strong>toAddress</strong> – your intended recipients of the report, separate multiple recipients with a comma with no spaces.
<strong>fromAddress</strong> – the primary smtp address of the Shared Mailbox or User Mailbox you want to use for sending the email report.
<strong>organizationName </strong>– the name of your organization to reflect in the alert.

<h3>
How to Exclude Workloads from the Report</h3>
<b>Note: This is applicable only from version 1.3.</b>
<ol>
<li>Open the <b><i>\resources\exclusions.csv</i></b> file</li>
<li>Change the Excluded value of the workload you want to exclude to 1 (0=include, 1=include)</li>
</ol>
<div class="separator" style="clear: both; text-align: center;">
<a href="https://2.bp.blogspot.com/-copZIVSmKas/XHLb1215aAI/AAAAAAAAFQI/Xfpy_gNaK1o_zMl4CH2a-o5_L1Wuc5w5wCLcBGAs/s1600/exclusions.csv.png" style="margin-left: 1em; margin-right: 1em;"><img border="0" data-original-height="428" data-original-width="474" src="https://2.bp.blogspot.com/-copZIVSmKas/XHLb1215aAI/AAAAAAAAFQI/Xfpy_gNaK1o_zMl4CH2a-o5_L1Wuc5w5wCLcBGAs/s1600/exclusions.csv.png" /></a></div>
<h3>
Running the Script</h3>
<strong>IMPORTANT</strong>: In the first run, whether in Test Mode or not, will only generate the data that will be needed for future run comparisons.
In this example, the script is in run Test Mode.
<div class="separator" style="clear: both; text-align: center;">
<a href="https://lh3.googleusercontent.com/-AJFlfyqFSM0/W_4gAyJO5DI/AAAAAAAAEKc/owBaw6V2eNM6hS-S-Wynee-CoRdu0mvmQCHMYCw/s1600-h/mRemoteNG_2018-11-28_12-49-03%255B4%255D" style="margin-left: 1em; margin-right: 1em;" target="_blank"><img alt="" border="0" height="204" src="https://lh3.googleusercontent.com/-NGL_t0MdABM/W_4gCrhgy4I/AAAAAAAAEKg/syB7f0qOGeofVu55aw_m_LZElVAOqd_4gCHMYCw/mRemoteNG_2018-11-28_12-49-03_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="531" /></a></div>
<h3>
Sample Output</h3>
<h4>
Email</h4>
<a href="https://lh3.googleusercontent.com/-ZazwhEsa4eY/W_4g5wLOmrI/AAAAAAAAELY/FXAKahTrNrwmI1o9rIF-97WRLdb1hm0owCHMYCw/s1600-h/mRemoteNG_2018-11-28_12-50-41%255B6%255D" target="_blank"><img alt="" border="0" height="821" src="https://lh3.googleusercontent.com/-OiUiWz14wVU/W_4g8Er-_PI/AAAAAAAAELc/XqK-QQo7GBYF7ma4BeDW3redGp3Hoc2agCHMYCw/mRemoteNG_2018-11-28_12-50-41_thumb%255B4%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="701" /></a>
<h4>
HTML</h4>
<a href="https://lh3.googleusercontent.com/-JW3OtGimoS4/W_4gIGlsQAI/AAAAAAAAEKs/2vawZp_f0lo6Q_pg8EW3DgMnKKU_ovMIQCHMYCw/s1600-h/mRemoteNG_2018-11-28_12-52-22%255B3%255D"><img alt="mRemoteNG_2018-11-28_12-52-22" border="0" height="177" src="https://lh3.googleusercontent.com/-rQ9dW8SqfGQ/W_4gJnProVI/AAAAAAAAEKw/PUiHvxUWl7YVuG-ORnddO0qt7Y3mFOPQwCHMYCw/mRemoteNG_2018-11-28_12-52-22_thumb%255B1%255D?imgmax=800" style="background-image: none; display: inline;" title="mRemoteNG_2018-11-28_12-52-22" width="401" /></a>
<a href="https://lh3.googleusercontent.com/-U2cKcnHEtzU/W_4g-31jCxI/AAAAAAAAELg/z708htU7cj00L5k7H_5U2umpxfaZ9WcbQCHMYCw/s1600-h/mRemoteNG_2018-11-28_12-53-17%255B4%255D" target="_blank"><img alt="" border="0" height="632" src="https://lh3.googleusercontent.com/-d2l3SqgZAME/W_4hBsi4HSI/AAAAAAAAELk/1KnUvxkvi7U-B6NrkUeQgq6hnsQ3A4RkACHMYCw/mRemoteNG_2018-11-28_12-53-17_thumb%255B2%255D?imgmax=800" style="background-image: none; display: inline;" title="" width="1038" /></a>

This script is functional, but I’m sure there can be many improvements. Or perhaps someone else has&nbsp;accomplished this differently. So please feel free to comment or modify and improve, just please don’t forget to credit the original source.