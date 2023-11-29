i) You need a "C:\scripts\ChatGPTanswersZendesk"  folder.
ii) https://runpcrun.zendesk.com/admin/apps-integrations/apis/zendesk-api/settings - add an API token, give it a suitable name, and copy it into your ITGlue personal password section.
iii) there is a "config.json" file, put your Zendesk API token in there and your zendesk email address in the appropriate place
iv) if you have placed the files in "C:\scripts\ChatGPTanswersZendesk" then just run the "install.reg" file. Otherwise edit it to the appropriate place.
v) Create a new bookmark for anywhere and name it something like "ZenGPT" or whatever you like. Open bookmarklet.txt and copy and paste the javascript into the URL.
vi) run it on a test ticket such as https://runpcrun.zendesk.com/agent/tickets/86448 - the first time it runs you will get a dialog and no choice about making it a default action.
The second time it will allow setting a default action.