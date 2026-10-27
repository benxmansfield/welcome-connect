// Welcome & Connect newcomer tracker: settings.
// Only clientId needs filling in, after you register the app in Microsoft Entra (see README.md).
window.WC_CONFIG = {
  // Application (client) ID from the Entra app registration
  clientId: "854b99dd-4be8-4b97-af8a-7bf1d65a2ebe",

  // Southside's Microsoft 365 tenant
  tenantId: "3b60f06f-e956-4133-897f-6a996928f315",

  // The Welcome & Connect SharePoint site (Graph site id: host,siteCollectionId,webId)
  siteId: "southsideau.sharepoint.com,aa789c83-b209-46fd-ac43-a585d8a26d7d,cbd18d11-a0ca-4382-90fc-f179f4a0a37d",
  siteUrl: "https://southsideau.sharepoint.com/sites/WelcomeConnect",

  // List names (must match SharePoint exactly)
  lists: { nextSteps: "Next Steps", people: "People", team: "Team" },

  // How many Sundays back the This Sunday screen can step through
  sundaysBack: 8,

  // Where to send people who aren't on the Team list
  leaderName: "Ben Mansfield"
};
