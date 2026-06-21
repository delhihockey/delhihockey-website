# Delhi Hockey — Official Website

The official website of Delhi Hockey, the recognised state association affiliated to Hockey India and the Delhi Olympic Association.

## How this site is hosted

- **Code:** stored on GitHub
- **Hosting:** Cloudflare Pages (free, unlimited bandwidth)
- **Database:** Supabase (stores documents, league data, sponsors, etc.)

## Cloudflare Pages build settings

When connecting this repository in Cloudflare Pages, use:

- **Framework preset:** Vite
- **Build command:** `npm run build`
- **Build output directory:** `dist`

## Managing content

Content is managed through the built-in admin panel:
1. Scroll to the footer of the live site
2. Click **Admin**
3. Enter the admin password
4. Manage documents, the league, sponsors, registration links, and more

## Changing the admin password

Open `src/DelhiHockey.jsx` and find the line near the top:

```js
const ADMIN_PASSWORD = "delhihockey2024";
```

Change the text in quotes, save, and commit. The site rebuilds automatically.
