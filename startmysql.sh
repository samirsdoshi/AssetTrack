CONTAINER=asset
docker stop asset
docker rm asset
docker run -d -p 3306:3306 \
 --name assets \
  -e MYSQL_ROOT_PASSWORD=sa123  \
  -default-authentication-plugin=mysql_native_password \
  --volume='/Users/Sdoshi/development/samir/docs/Investments/US/data':/var/lib/mysql \
  mysql:latest
  
