web: gunicorn app:app --workers 2 --threads 2 --worker-class gthread --timeout 120 --bind 0.0.0.0:$PORT --access-logfile - --error-logfile -
