# Data_Validation
compare 2 files with duplicates and some basic DQ


How others use it
Locally: they clone the repo and double-click start.bat, or run ./start.sh, then open http://localhost:5000.
Docker: docker build -t mrna-validator . followed by docker run -p 5000:5000 mrna-validator.
Shared link: GitHub itself can't run a Flask backend, so a link for others needs a host such as Render, Railway or Azure. Point it at the Dockerfile or use gunicorn app:app.
