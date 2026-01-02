FROM mysql:5.6
# Add the content of the sql-scripts/ directory to your image
# All scripts in docker-entrypoint-initdb.d/ are automatically
# executed during container startup
#COPY ./scripts/init /docker-entrypoint-initdb.d/
