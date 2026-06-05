# Radio Room Blog Publisher API

Flask API that converts a .docx blog post to HTML and publishes it to SiteGround via FTP.

## Deploy to Render

1. Push this `render-api/` folder to a GitHub repo
2. In Render: New → Web Service → connect the repo
3. Set Build Command: `pip install -r requirements.txt`
4. Set Start Command: `gunicorn app:app`
5. Add these Environment Variables in Render dashboard:

| Variable       | Value                              |
|----------------|------------------------------------|
| `FTP_HOST`     | Your SiteGround FTP hostname       |
| `FTP_USER`     | FTP username                       |
| `FTP_PASS`     | FTP password                       |
| `FTP_BLOG_PATH`| `/public_html/blog`                |
| `API_KEY`      | Any secret string you choose       |

## Test it

Once deployed, test with curl:

```bash
curl -X POST https://your-app.onrender.com/convert \
  -H "X-API-Key: your-secret-key" \
  -F "file=@my-post.docx"
```

Without FTP configured it returns the HTML in the response — useful for testing.

## Make scenario

1. Watch Google Drive folder for new files
2. Download the file
3. HTTP POST to `https://your-app.onrender.com/convert`
   - Header: `X-API-Key: your-secret-key`
   - Body: multipart, field name `file`, content = the .docx
4. Done — post is live on SiteGround
