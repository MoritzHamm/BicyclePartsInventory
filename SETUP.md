# Bike Parts Inventory — setup

Two pieces: an **Apps Script backend** (lives in your Google Sheet) and a **scanning page**
(a normal web page you host so the camera works).

## 1. Backend (5 min)

1. Create a new Google Sheet.
2. **Extensions → Apps Script**. Delete the sample code, paste all of `Code.gs`.
3. Change `SECRET` (line 22) to a long random string. Remember it.
4. **Run → `setup`**. Approve the permission prompt. It builds the `Products`, `Items`,
   `Bikes` tabs, the reference formulas, and a starter bike.
5. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy, then **copy the Web app `/exec` URL**.

> Re-deploying later: use **Manage deployments → edit → new version**, so the URL stays the same.

## 2. Scanner page (5 min)

1. Open `index.html`, set the two config lines near the bottom:
   ```js
   const API_URL = "…your /exec URL…";
   const SECRET  = "…same secret as in Code.gs…";
   ```
2. Host `index.html` + `manifest.webmanifest` over **https** (camera needs a secure origin):
   - **GitHub Pages** (drop both files in a repo, enable Pages), or
   - **Cloudflare Pages**, or
   - your own box with a cert.
3. Open the URL on your phone → **Start camera** → allow camera access.
4. Optional: browser menu → **Add to Home screen** for an app-like icon.

## Using it

- **Scan tab** = intake. Every scan makes one physical item in `storage` and shows a number —
  write that number on the part (paint pen / numbered zip tie for greasy chains). Scan the same
  box twice to register two items.
- **No barcode?** On the Scan tab, **Add without barcode** lets you pick an existing product or
  create a new one, with a quantity. Barcode-less products get a synthetic key like `NB-0001`
  and otherwise behave exactly like scanned ones — so your whole bottle-cage pile is one product
  entry plus N item instances. Each still gets its own number to mark.
- **Products tab** in the Sheet: fill in `brand` / `product_name` / `category` for each barcode
  **once**. It auto-fills across every item with that barcode.
- **Items tab** = the Manage tab in the app: tap an item to **Install** (pick bike + odometer),
  **Remove** (back to storage, km rolls up), or **Retire**.
- **Bikes tab**: add your bikes and keep `odometer_km` current. Lifetime km per item is computed
  from the difference between install and removal odometer readings. (This is where a future
  Strava sync would just write the odometer numbers for you.)

## Notes

- The `current_km`, `brand`, `product_name` columns in `Items` are live formulas — don't type over them.
- `SECRET` keeps random people who find the URL from writing to your sheet. It's light protection,
  fine for a personal tool; don't put anything sensitive in the sheet.
