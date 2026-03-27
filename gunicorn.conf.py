import os


# Large XLS files can take a while to rebuild on Render's small instances.
# Give the request enough time to finish instead of letting Gunicorn kill it.
timeout = int(os.environ.get("GUNICORN_TIMEOUT", "300"))
graceful_timeout = int(os.environ.get("GUNICORN_GRACEFUL_TIMEOUT", "30"))
workers = int(os.environ.get("WEB_CONCURRENCY", "1"))
