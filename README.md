# CSST-106 Computer Vision Lab

A static web interface for browsing seven Computer Vision lab handouts. Selecting **View handout** displays the original DOCX in a browser-based reader with zoom controls.

## Run locally

From the project directory, run:

```powershell
python server.py
```

Then open `http://localhost:4173/`.

## GitHub Pages

Every push to `main` deploys the site through the workflow in `.github/workflows/pages.yml`.

After deployment, the site is available at <https://inigss.github.io/CSST-106/>.
