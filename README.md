# Welcome & Connect: newcomer tracker (version 1)

A small phone-first web app for the Southside welcome team. It reads and writes the lists in the
**Welcome & Connect** SharePoint site (`https://southsideau.sharepoint.com/sites/WelcomeConnect`)
using Microsoft Graph. Team members sign in with their church Microsoft 365 account, or as a guest
with any email address (Microsoft sends them a one-time code).

Screens in version 1:

- **This Sunday** (Leaders and Team leaders): new people, how many were triaged, by service, and
  the cards still waiting for triage. Arrows step back through the last 8 Sundays.
- **Add new person**: the Next steps form. Saves one row to the *Next Steps* list.
- **Your recent cards** (Connectors): their own recent cards. SharePoint only lets connectors see their own.
- **Not on the team yet**: shown to anyone signed in who isn't on the *Team* list.

There is no server. The app is plain HTML, CSS and JavaScript. Microsoft's sign-in library (MSAL) loads from jsDelivr, pinned to 3.27.0 with an integrity hash.

## 1. Register the app in Microsoft Entra (about 10 minutes, once)

You need to be a Microsoft 365 admin for this.

1. Go to **entra.microsoft.com** › Identity › Applications › **App registrations** › **New registration**.
2. Name: `Welcome & Connect tracker`. Supported account types: **this organisational directory only**.
3. Redirect URI: platform **Single-page application (SPA)**, value = the address where the app will live,
   for example `https://benxmansfield.github.io/welcome-connect/` (keep the trailing slash).
   Add `http://localhost:8080/` too if you want to test on your computer.
4. Click **Register**. Copy the **Application (client) ID**.
5. **API permissions** › Add a permission › Microsoft Graph › **Delegated** › add `User.Read` and
   `Sites.ReadWrite.All`. Then click **Grant admin consent for Southside**.
6. Paste the client ID into `config.js` (`clientId`).

`Sites.ReadWrite.All` is *delegated*: the app can only ever do what the signed-in person could
already do in SharePoint. The SharePoint permissions below are what keep data private.

## 2. Put it online (GitHub Pages)

1. Create a new public repository, e.g. `welcome-connect`, and upload these files.
2. Settings › Pages › Deploy from branch › `main` / root.
3. The app is at `https://<your-github-name>.github.io/welcome-connect/`. That must match the redirect URI in step 1.

The repository only holds code. No names or notes about newcomers are ever stored in GitHub.

To test locally instead: run `python3 -m http.server 8080` in this folder and open `http://localhost:8080/`.

## 3. Give the team access

Two things are needed for each person:

1. **A row in the Team list** with their name, church email, Role (`Leader`, `Team leader` or `Connector`),
   service, and Active = Yes.
2. **Access to the SharePoint site** at the right level:

| Role | SharePoint access | Why |
| --- | --- | --- |
| Leader | Site **Owners** group | Sees everything, including Care notes |
| Team leader | Site **Members** group (Edit) | Edit includes "Manage lists", so they see every card |
| Connector | A **Connectors** group with **Contribute** | Next Steps has item-level permissions, so they only see their own cards |

To make the Connectors group: Site settings › Site permissions › Advanced permissions settings ›
Create group › name `Welcome & Connect Connectors` › permission level **Contribute**.
Also stop connectors seeing the People and Team lists: in each of those lists, go to
List settings › Permissions for this list › Stop inheriting permissions, and remove the Connectors group.

## 4. Share it

Send the team the app's link and ask them to add it to their phone's home screen
(iPhone: Share › Add to Home Screen; Android: menu › Add to Home screen).

## How the data is stored

| App | SharePoint list and column |
| --- | --- |
| Their name / description | Next Steps › Name or description, Has name |
| Which service? | Next Steps › Service |
| First time? | Next Steps › Visit |
| What did you learn? | Next Steps › What they learned (Note) |
| Who will follow up? | Next Steps › Follow-up (`I will` / `Welcome team`) |
| Where are they at? | Next Steps › Road (`Not sure yet` until triaged) |
| Interested in | Next Steps › Interested in |
| Contact, preference | Next Steps › Contact, Contact preference |
| Sunday | Next Steps › Sunday (the most recent Sunday, Brisbane time) |
| (automatic) | Next Steps › Source = `Team card`, Matched = No, Created By = who filled it in |

"Triaged" in the app means a card whose Road is anything other than `Not sure yet`.
A team leader can update a card's Road, or tick Matched, by opening it in SharePoint ("Open" in the app).

## What's not in version 1 yet

- The People list screens (matching cards to people, owners, follow-up dates)
- The quarter overview
- Copying Microsoft Form answers from the trial into Next Steps (a Power Automate flow can do this)
- The Tuesday reminder email (Power Automate, reading the People list)
