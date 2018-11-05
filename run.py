#!flask/bin/python

from app import app

app.run(threaded=True, port=5000, debug=True)
