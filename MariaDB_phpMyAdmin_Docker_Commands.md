
# MariaDB 10.5.29 + phpMyAdmin Setup using Docker

## 1. Install Docker
```bash
sudo apt update
sudo apt install docker.io -y
sudo systemctl enable --now docker
````

---

## 2. Create a Docker Volume for MariaDB

```bash
docker volume create mariadb10529_data
```

---

## 3. Run MariaDB 10.5.29 Container

```bash
docker run -d \
  --name mariadb10529 \
  -e MARIADB_ROOT_PASSWORD=cyber@root123 \
  -e MARIADB_DATABASE=nadalbus_bgv \
  -e MARIADB_USER=nadalbus_bgv \
  -e MARIADB_PASSWORD=Nadal@bgv123 \
  -v mariadb10529_data:/var/lib/mysql \
  -p 3306:3306 \
  mariadb:10.5.29
```

**Access MariaDB CLI inside the container:**

```bash
docker exec -it mariadb10529 mariadb -u root -p
```

---

## 4. Create phpMyAdmin Container

```bash
docker run -d \
  --name phpmyadmin \
  --link mariadb10529:db \
  -p 8080:80 \
  -e PMA_HOST=db \
  -e PMA_PORT=3306 \
  phpmyadmin/phpmyadmin
```

**Access phpMyAdmin in browser:**

```
http://<your-server-ip>:8080
```

Use the database credentials set above.

---

## 5. Docker Basic Commands

**List all running containers:**

```bash
sudo docker ps
```

**Start a container:**

```bash
sudo docker start mariadb10529
```

**Stop a container:**

```bash
sudo docker stop mariadb10529
```

**Check container logs:**

```bash
sudo docker logs mariadb10529
```

# [OR] Run Auto mode
Create a yml file named `docker-compose.yml`
```yml

version: '3.8'

services:
  mariadb:
    image: mariadb:10.5
    container_name: mariadb10529
    environment:
      MYSQL_ROOT_PASSWORD: rootpass
      MYSQL_DATABASE: mydb
      MYSQL_USER: myuser
      MYSQL_PASSWORD: mypass
    ports:
      - "3306:3306"
    volumes:
      - mariadb_data:/var/lib/mysql

  phpmyadmin:
    image: phpmyadmin/phpmyadmin
    container_name: phpmyadmin
    environment:
      PMA_HOST: mariadb
      PMA_PORT: 3306
    ports:
      - "8080:80"
    depends_on:
      - mariadb

volumes:
  mariadb_data:
```
Restart docker
```bash
sudo systemctl restart docker
```
Up the Container
```bash
sudo docker-compose up -d
```
See the services running 
```bash
sudo docker ps
```
view all 
```bash
sudo docker ps -a
```
Down all
Up the Container
```bash
sudo docker-compose down
```
for more see running ports 
```bash
sudo ss -tulpn | grep 3306
```
### Runn a Big SQL File in single short 
Copy the SQL file
```bash
sudo docker cp 'file_path/file_name.sql'  mariadb10529:/tmp/sql_query.sql
```
Run the SQL file
```bash
sudo docker exec -it mariadb10529 bash -c "mariadb -u root -prootpass nadalbus_bgv < /tmp/sql_Query.sql"
```



