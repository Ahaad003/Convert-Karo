# ConvertKaro — Free File Converter

PNG/JPG image convert, PDF se image, PDF↔Word, aur PDF editing — sab free.

## Project Structure

```
convert-karo/
├── public/
│   └── index.html      ← Frontend (poori website ka UI)
├── server/
│   └── index.js         ← Backend (Node.js/Express — conversions handle karta hai)
├── package.json          ← Dependencies list
└── README.md
```

## Website Live Kaise Karein (Step-by-Step)

### Step 1: GitHub par Upload karo

1. GitHub.com par jao, login karo (account: Ahad003)
2. New repository banao — naam: `convert-karo`, Public rakho
3. Is poore `convert-karo` folder ke andar ke saare files/folders upload karo
   (GitHub website par "uploading an existing file" option se drag-drop kar sakte ho,
   ya `git` command line se: `git add .` → `git commit -m "Initial ConvertKaro"` → `git push`)

### Step 2: Railway par Deploy karo

1. railway.app par jao, GitHub se login karo
2. "New Project" → "Deploy from GitHub repo" → `convert-karo` select karo
3. Railway apne aap `package.json` dekh ke Node.js app detect kar lega
4. **IMPORTANT**: Railway ke settings mein ek **Nixpacks/Build setting** add karni hogi taaki
   LibreOffice bhi server par install ho (PDF↔Word ke liye zaroori hai) —
   iske liye root mein ek `nixpacks.toml` file chahiye (neeche di gayi hai)
5. Deploy hone do (5-10 min lag sakte hain LibreOffice install hone mein)
6. "Settings" → "Networking" → "Generate Domain" — yahan se public URL milega
   (jaise `convert-karo-production.up.railway.app`)

### Step 3: Custom Domain (optional, baad mein)

1. Namecheap/GoDaddy se `convertkaro.com` jaisa domain khareedo (~₹700-1000/year)
2. Railway settings mein "Custom Domain" add karo, jo DNS records milein wo domain
   provider ki settings mein daal do
3. 24-48 ghante mein live ho jayega custom domain par

## Features Status

| Feature | Status |
|---|---|
| PNG ↔ JPG ↔ WEBP convert | ✅ Working (frontend hi, backend ki zaroorat nahi) |
| PDF → Image | ✅ Working (frontend hi) |
| PDF → Word | ✅ Backend ready (Railway par LibreOffice install hone ke baad live hoga) |
| Word → PDF | ✅ Backend ready |
| PDF Editor (text correction) | ✅ Basic version ready (white-box overlay method) |

## Next Steps (baad mein add karna)

- [ ] Google Analytics tracking code (`<head>` mein GA4 script daalna)
- [ ] Google Search Console verification
- [ ] Wallet/payment system (agar paid features add karne hain)
- [ ] File size limits (abuse rokne ke liye)
- [ ] Rate limiting (ek user zyada requests na bhej sake)
