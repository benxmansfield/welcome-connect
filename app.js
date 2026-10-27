/* Welcome & Connect newcomer tracker, version 1.
   Screens: sign in, not on team, This Sunday (leaders and team leaders),
   Add new person (Next steps), My cards (connectors).
   Data lives in the Welcome & Connect SharePoint site; read and written with Microsoft Graph
   as the signed-in person, so SharePoint permissions decide what each person can see. */
(function () {
  "use strict";
  const C = window.WC_CONFIG;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["User.Read", "Sites.ReadWrite.All"];
  const TZ = "Australia/Brisbane";
  const app = document.getElementById("app");
  const nav = document.getElementById("nav");
  const meBtn = document.getElementById("me");

  const ROADS = ["Still figuring it out", "Coming back to faith", "I know Jesus", "Not sure yet"];
  const INTERESTS = ["Discover", "Have a chat", "Growth Groups", "Serving Teams", "Kids and youth"];
  const INTEREST_LABELS = { "Have a chat": "A chat", "Serving Teams": "Serving" };

  const state = { me: null, role: null, listIds: {}, sundayIdx: 0, cards: null };

  // ---------- small helpers ----------
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const ymd = (d) => new Intl.DateTimeFormat("en-CA", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit" }).format(d);
  const prettyDate = (s) => new Intl.DateTimeFormat("en-AU", { timeZone: TZ, day: "numeric", month: "long" }).format(new Date(s + "T12:00:00+10:00"));
  const brisMidnightUtc = (s) => new Date(s + "T00:00:00+10:00").toISOString(); // Brisbane has no daylight saving
  const fieldDate = (v) => (v ? ymd(new Date(v)) : "");
  function lastSundays(n) {
    const today = ymd(new Date());
    const d = new Date(today + "T12:00:00+10:00");
    const dow = new Intl.DateTimeFormat("en-AU", { timeZone: TZ, weekday: "short" }).format(d);
    const back = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].indexOf(dow);
    const out = [];
    for (let i = 0; i < n; i++) out.push(ymd(new Date(d.getTime() - (back + 7 * i) * 864e5)));
    return out; // newest first
  }
  const chevron = (dir) => `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${dir === "l" ? "M15 6l-6 6 6 6" : "M9 6l6 6-6 6"}"/></svg>`;
  const render = (html) => { app.innerHTML = html; window.scrollTo(0, 0); };
  const fail = (e) => {
    console.error(e);
    render(`<div class="card"><h2>Something went wrong</h2><p class="muted" style="margin-top:8px">${esc(e.message || e)}</p><button class="btn btn-ghost" style="margin-top:14px" onclick="location.reload()">Try again</button></div>`);
  };

  // ---------- sign-in ----------
  const msalApp = new msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: "https://login.microsoftonline.com/" + C.tenantId, redirectUri: location.origin + location.pathname },
    cache: { cacheLocation: "localStorage" }
  });

  async function token() {
    const account = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!account) throw new Error("Not signed in");
    try {
      return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
    } catch (e) {
      await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
      throw e;
    }
  }

  async function graph(path, opts = {}) {
    const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
      ...opts,
      headers: {
        Authorization: "Bearer " + (await token()),
        "Content-Type": "application/json",
        Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
        ...(opts.headers || {})
      }
    });
    if (!res.ok) {
      const t = await res.text();
      const err = new Error(res.status === 403 ? "You don't have access to the Welcome & Connect site yet. Ask " + C.leaderName + " to add you." : "Microsoft 365 said: " + res.status + " " + t.slice(0, 200));
      err.status = res.status;
      throw err;
    }
    return res.status === 204 ? null : res.json();
  }

  async function allItems(listId, query) {
    let url = `/sites/${C.siteId}/lists/${listId}/items?${query}`;
    const out = [];
    while (url) {
      const page = await graph(url);
      out.push(...page.value);
      url = page["@odata.nextLink"] || null;
    }
    return out;
  }

  async function loadContext() {
    const me = await graph("/me?$select=displayName,mail,userPrincipalName");
    state.me = me;
    const lists = await graph(`/sites/${C.siteId}/lists?$select=id,displayName`);
    for (const l of lists.value) state.listIds[l.displayName] = l.id;
    const team = await allItems(state.listIds[C.lists.team], "expand=fields(select=Title,Email,Role,Active)&$top=200");
    const emails = [me.mail, me.userPrincipalName].filter(Boolean).map((s) => s.toLowerCase());
    const row = team.find((t) => t.fields.Active !== false && emails.includes(String(t.fields.Email || "").toLowerCase()));
    state.role = row ? row.fields.Role || "Connector" : null;
    meBtn.textContent = (me.displayName || "?").trim().charAt(0).toUpperCase();
    nav.hidden = false;
    nav.querySelectorAll('[data-role="lead"]').forEach((a) => (a.hidden = !isLead()));
  }
  const isLead = () => state.role === "Leader" || state.role === "Team leader";

  meBtn.addEventListener("click", () => {
    if (confirmSignOut()) msalApp.logoutRedirect({ account: msalApp.getActiveAccount() });
  });
  function confirmSignOut() {
    // A simple in-page confirm so we don't rely on browser dialogs.
    render(`<div class="card center"><h2>Sign out?</h2><p class="muted" style="margin:8px 0 16px">You'll need your email to sign back in.</p>
      <div style="display:flex;gap:8px"><button id="so-yes" class="btn btn-primary" style="flex:1">Sign out</button><a href="#/" class="btn btn-ghost" style="flex:1">Cancel</a></div></div>`);
    document.getElementById("so-yes").onclick = () => msalApp.logoutRedirect({ account: msalApp.getActiveAccount() });
    return false;
  }

  // ---------- screens ----------
  function screenSignIn() {
    nav.hidden = true;
    render(`<div class="card center" style="padding:28px 20px">
      <h1>Welcome &amp; Connect</h1>
      <p class="muted" style="margin:8px 0 20px">For the Southside welcome team. Sign in with the email address you were invited with.</p>
      <button id="signin" class="btn btn-primary" style="width:100%">Sign in</button>
    </div>`);
    document.getElementById("signin").onclick = () => msalApp.loginRedirect({ scopes: SCOPES });
  }

  function screenNotOnTeam() {
    render(`<div class="card">
      <h1>You're not on the team list yet</h1>
      <p style="margin-top:10px">You're signed in as <strong>${esc(state.me.mail || state.me.userPrincipalName)}</strong>, but this account hasn't been added to the welcome team.</p>
      <p style="margin-top:10px">Ask ${esc(C.leaderName)} to add you, then open this page again.</p>
    </div>`);
  }

  async function loadCards() {
    const since = lastSundays(C.sundaysBack).at(-1);
    const items = await allItems(
      state.listIds[C.lists.nextSteps],
      "expand=fields(select=Title,HasName,Sunday,Service,Visit,Note,Road,FollowUp,Matched,Source)" +
        `&$filter=fields/Sunday ge '${brisMidnightUtc(since)}'&$top=500`
    );
    items.sort((a, b) => String(b.createdDateTime).localeCompare(String(a.createdDateTime)));
    state.cards = items.map((i) => ({ id: i.id, created: i.createdDateTime, by: i.createdBy?.user?.displayName || "", f: i.fields, sunday: fieldDate(i.fields.Sunday) }));
  }

  async function screenSunday() {
    if (!isLead()) return (location.hash = "#/mine");
    render(`<p class="muted center pad">Loading this Sunday…</p>`);
    if (!state.cards) await loadCards();
    const sundays = lastSundays(C.sundaysBack);
    const i = Math.min(state.sundayIdx, sundays.length - 1);
    const day = sundays[i];
    const firsts = state.cards.filter((c) => c.sunday === day && c.f.Visit === "First time");
    const tri = firsts.filter((c) => c.f.Road && c.f.Road !== "Not sure yet");
    const by = (svc) => { const n = firsts.filter((c) => c.f.Service === svc); return `${n.filter((c) => c.f.Road && c.f.Road !== "Not sure yet").length} of ${n.length}`; };
    const pct = firsts.length ? Math.round((tri.length / firsts.length) * 100) + "%" : "–";
    const dots = firsts.map((c) => `<span class="dot ${c.f.Road && c.f.Road !== "Not sure yet" ? "tri" : "un"}"></span>`).join("");
    const todo = state.cards.filter((c) => (!c.f.Road || c.f.Road === "Not sure yet") && !c.f.Matched).slice(0, 5);

    render(`
      <section class="card" aria-labelledby="sun-h">
        <div class="sun-head">
          <button class="navbtn" id="prev" aria-label="Previous Sunday" ${i >= sundays.length - 1 ? "disabled" : ""}>${chevron("l")}</button>
          <div class="sun-title"><h1 id="sun-h">${i === 0 ? "This Sunday · " : "Sunday "}${esc(prettyDate(day))}</h1></div>
          <button class="navbtn" id="next" aria-label="Next Sunday" ${i === 0 ? "disabled" : ""}>${chevron("r")}</button>
        </div>
        <div class="big"><div><strong>${firsts.length}</strong><span>new people</span></div><div><strong>${tri.length}</strong><span>triaged</span></div></div>
        <div class="dots" role="img" aria-label="${tri.length} of ${firsts.length} new people triaged">${dots || '<span class="muted">No cards yet for this Sunday.</span>'}</div>
        <div class="split"><span><strong>${pct}</strong> triaged</span><span>8:30 · ${by("8:30")}</span><span>10:30 · ${by("10:30")}</span></div>
      </section>
      <a class="btn btn-primary" href="#/add">＋ Add new person</a>
      <section>
        <div class="row-head"><h2>Triage these first</h2><a href="${esc(C.siteUrl)}/Lists/Next%20Steps/AllItems.aspx" target="_blank" rel="noopener" style="font-size:14px">All cards</a></div>
        <ul class="list" style="margin-top:10px">${todo.length ? todo.map(cardRow).join("") : '<li class="item"><span class="muted">Nothing waiting. Nice work.</span></li>'}</ul>
      </section>`);
    document.getElementById("prev").onclick = () => { state.sundayIdx = i + 1; screenSunday(); };
    document.getElementById("next").onclick = () => { state.sundayIdx = Math.max(0, i - 1); screenSunday(); };
  }

  function cardRow(c) {
    const tag = c.f.FollowUp === "Welcome team" ? '<span class="tag tag-team">Welcome team</span>' : c.f.Visit === "First time" ? '<span class="tag tag-new">First time</span>' : "";
    const meta = [c.f.Service, c.sunday && prettyDate(c.sunday), c.by && "by " + c.by].filter(Boolean).join(" · ");
    return `<li class="item"><div class="grow"><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><span class="name">${esc(c.f.Title)}</span>${tag}</div>
      <span class="meta">${esc(meta)}</span></div>
      <a class="btn btn-ghost btn-sm" href="${esc(C.siteUrl)}/Lists/Next%20Steps/DispForm.aspx?ID=${encodeURIComponent(c.id)}" target="_blank" rel="noopener">Open</a></li>`;
  }

  async function screenMine() {
    render(`<p class="muted center pad">Loading…</p>`);
    if (!state.cards) await loadCards();
    const mine = state.cards.slice(0, 10); // SharePoint already limits connectors to their own cards
    render(`
      <a class="btn btn-primary" href="#/add">＋ Add new person</a>
      <section>
        <div class="row-head"><h2>${isLead() ? "Recent cards" : "Your recent cards"}</h2></div>
        <ul class="list" style="margin-top:10px">${mine.length ? mine.map(cardRow).join("") : '<li class="item"><span class="muted">No cards yet.</span></li>'}</ul>
      </section>`);
  }

  // ---------- Add new person (Next steps) ----------
  function pills(name, options, multi, labels = {}) {
    return `<div class="pills" role="group" data-name="${name}" data-multi="${multi ? 1 : 0}">` +
      options.map((o) => `<button type="button" class="pill" aria-pressed="false" data-value="${esc(o)}">${esc(labels[o] || o)}</button>`).join("") + "</div>";
  }

  function screenAdd() {
    const sundays = lastSundays(1);
    render(`
      <form id="f" class="card" style="display:flex;flex-direction:column;gap:20px" novalidate>
        <div><h1>Next steps</h1><p class="muted" style="margin-top:4px">Met someone new? The first five are all we need.</p></div>

        <div class="q" id="q-name">
          <label class="lbl" for="name">1. Their name</label>
          <input id="name" class="field" autocomplete="off" placeholder="First name is fine">
          <button type="button" id="toggle-name" class="linkbtn">Didn't catch their name?</button>
          <span id="desc-hint" class="hint" hidden>No looks. Where they sat, who they came with, what you chatted about.</span>
        </div>

        <div class="q"><span class="lbl">2. Which service?</span>
          <div class="seg">${pills("service", ["8:30", "10:30"], false).replace('<div class="pills"', '<div class="pills" style="display:contents"')}</div></div>

        <div class="q"><span class="lbl">3. Is this their first time?</span>${pills("visit", ["First time", "Been before", "Not sure"], false)}</div>

        <div class="q"><label class="lbl" for="note">4. What did you learn?</label>
          <textarea id="note" class="field" placeholder="e.g. moved to Brisbane in September, grew up in church"></textarea>
          <span class="hint">Write it as if they might read it one day.</span></div>

        <div class="q"><span class="lbl">5. Who will follow up?</span>${pills("follow", ["I will", "Welcome team"], false, { "Welcome team": "Welcome team, please" })}</div>

        <details class="more">
          <summary>Anything else? <span class="opt">(optional)</span><span aria-hidden="true">▾</span></summary>
          <div style="display:flex;flex-direction:column;gap:18px">
            <div class="q"><span class="lbl">Where are they at with Jesus?</span>${pills("road", ROADS, false)}<span class="hint">Only if it came up naturally. Picking one counts as triaged.</span></div>
            <div class="q"><span class="lbl">Interested in</span>${pills("interests", INTERESTS, true, INTEREST_LABELS)}</div>
            <div class="q"><label class="lbl" for="contact">Contact <span class="opt">(only if they offered it)</span></label>
              <input id="contact" class="field" autocomplete="off" placeholder="Mobile or email">${pills("pref", ["Text", "Call", "Email"], false)}</div>
          </div>
        </details>

        <p class="notice">Under 18? Please give a parent's name instead.</p>
        <p id="err" class="error" role="alert" hidden></p>
        <button class="btn btn-primary" id="send" type="submit">Send to the welcome team</button>
      </form>`);

    let noName = false;
    const nameEl = document.getElementById("name");
    document.getElementById("toggle-name").onclick = (e) => {
      noName = !noName;
      document.querySelector('label[for="name"]').textContent = noName ? "1. Describe them" : "1. Their name";
      nameEl.placeholder = noName ? "e.g. couple in their 60s, sat near the back" : "First name is fine";
      document.getElementById("desc-hint").hidden = !noName;
      e.target.textContent = noName ? "I know their name" : "Didn't catch their name?";
    };
    app.querySelectorAll(".pills").forEach((g) => {
      g.addEventListener("click", (e) => {
        const b = e.target.closest(".pill");
        if (!b) return;
        if (g.dataset.multi === "1") b.setAttribute("aria-pressed", b.getAttribute("aria-pressed") === "true" ? "false" : "true");
        else g.querySelectorAll(".pill").forEach((x) => x.setAttribute("aria-pressed", x === b && x.getAttribute("aria-pressed") !== "true" ? "true" : "false"));
      });
    });
    const picked = (name) => [...app.querySelectorAll(`[data-name="${name}"] .pill[aria-pressed="true"]`)].map((b) => b.dataset.value);

    document.getElementById("f").onsubmit = async (e) => {
      e.preventDefault();
      const err = document.getElementById("err");
      const title = nameEl.value.trim();
      const svc = picked("service")[0], visit = picked("visit")[0], follow = picked("follow")[0];
      const note = document.getElementById("note").value.trim();
      const missing = [!title && (noName ? "a description" : "their name"), !svc && "the service", !visit && "whether it's their first time", !note && "what you learned", !follow && "who will follow up"].filter(Boolean);
      if (missing.length) { err.hidden = false; err.textContent = "Please add " + missing.join(", ") + "."; return; }
      err.hidden = true;
      const btn = document.getElementById("send");
      btn.disabled = true; btn.textContent = "Sending…";
      const interests = picked("interests");
      const fields = {
        Title: title.slice(0, 255), HasName: !noName, Sunday: brisMidnightUtc(sundays[0]), Service: svc, Visit: visit, Note: note,
        Road: picked("road")[0] || "Not sure yet", FollowUp: follow, Source: "Team card", Matched: false
      };
      const contact = document.getElementById("contact").value.trim();
      if (contact) fields.Contact = contact.slice(0, 255);
      if (picked("pref")[0]) fields.ContactPreference = picked("pref")[0];
      if (interests.length) { fields["InterestedIn@odata.type"] = "Collection(Edm.String)"; fields.InterestedIn = interests; }
      try {
        await graph(`/sites/${C.siteId}/lists/${state.listIds[C.lists.nextSteps]}/items`, { method: "POST", body: JSON.stringify({ fields }) });
        state.cards = null;
        render(`<div class="card center" style="display:flex;flex-direction:column;gap:14px;padding:28px 20px">
          <div class="ok"><svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
          <h1>Thanks${state.me.displayName ? ", " + esc(state.me.displayName.split(" ")[0]) : ""}</h1>
          <p>${esc(title)} is on the list. ${follow === "I will" ? "Please follow up by Tuesday." : "A team leader will pick it up."}</p>
          <a class="btn btn-primary" href="#/add" id="again">Add someone else</a>
          <a class="btn btn-ghost" href="#/">Done</a></div>`);
        document.getElementById("again").onclick = (ev) => { ev.preventDefault(); screenAdd(); };
      } catch (ex) {
        btn.disabled = false; btn.textContent = "Send to the welcome team";
        err.hidden = false; err.textContent = ex.message;
      }
    };
  }

  // ---------- routing ----------
  async function route() {
    try {
      if (!msalApp.getActiveAccount()) return screenSignIn();
      if (!state.me) await loadContext();
      if (!state.role) return screenNotOnTeam();
      const h = location.hash || "#/";
      if (h.startsWith("#/add")) return screenAdd();
      if (h.startsWith("#/mine")) return await screenMine();
      if (h.startsWith("#/sunday") || h === "#/") return isLead() ? await screenSunday() : await screenMine();
      return isLead() ? await screenSunday() : await screenMine();
    } catch (e) {
      if (e.message !== "Not signed in") fail(e);
    }
  }

  (async function start() {
    if (!C.clientId || C.clientId.startsWith("PASTE")) {
      render(`<div class="card"><h2>Almost there</h2><p class="muted" style="margin-top:8px">Add the Application (client) ID to config.js. See README.md.</p></div>`);
      return;
    }
    try {
      await msalApp.initialize();
      const result = await msalApp.handleRedirectPromise();
      if (result?.account) msalApp.setActiveAccount(result.account);
      else if (!msalApp.getActiveAccount() && msalApp.getAllAccounts()[0]) msalApp.setActiveAccount(msalApp.getAllAccounts()[0]);
    } catch (e) { return fail(e); }
    window.addEventListener("hashchange", route);
    route();
  })();
})();
