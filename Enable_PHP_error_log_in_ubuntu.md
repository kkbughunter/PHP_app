# Enable Apache2 Error Log in LINUX
```bash
sudo tail -f /var/log/apache2/error.log
```
# Enable PHP Error Log in LINUX
## 1. Find Path
1. Create a test file in your website folder:

   ```bash
   echo "<?php phpinfo(); ?>" | sudo tee /var/www/html/phpinfo.php
   ```

2. Open it in your browser:

   ```
   http://localhost/phpinfo.php
   ```

3. Look for:

   ```
   Loaded Configuration File
   ```

   That’s the php.ini file your web PHP is using.


```
/etc/php/8.3/apache2/php.ini
```
## 2 Edit php.ini file to enable logging
4. Now you can enable logging for `/var/log/php_errors.log`.

---

### **1. Edit php.ini**

```bash
sudo nano /etc/php/8.3/apache2/php.ini
```

5. Find and set:

```ini
log_errors = On
error_log = /var/log/php_errors.log
```

Remove the `;` at the start if it’s commented out.

---

### **2. Create the log file & set permissions**

```bash
sudo touch /var/log/php_errors.log
sudo chown www-data:www-data /var/log/php_errors.log
sudo chmod 664 /var/log/php_errors.log
```

---

### **3. Restart Apache**

```bash
sudo systemctl restart apache2
```

---

### **4. Watch the log**

```bash
tail -f /var/log/php_errors.log
```

Now any PHP error from your web application will show here in real time.

---

### **5. run in background and clear and reopen 
`ctrl+z` for running in background 
```bash
clear && sudo truncate -s 0 /var/log/php_errors.log && fg
```

# Enable MySQL Error Log in LINUX
```
log_error = /var/log/mysql/error.log
```

inside the `[mysqld]` section of your MySQL or MariaDB server configuration file.

On Ubuntu 24.04, the typical main config files to check are:

* `/etc/mysql/mysql.conf.d/mysqld.cnf`
* `/etc/mysql/my.cnf` (sometimes just includes others)
* or any file inside `/etc/mysql/mysql.conf.d/`

**Steps:**

1. Open the main mysqld config file:

```bash
sudo nano /etc/mysql/mysql.conf.d/mysqld.cnf
```

2. Find the `[mysqld]` section (it should be near the top).

3. Add the line below inside that section:

```
log_error = /var/log/mysql/error.log
```

4. Save and exit.

5. Make sure the directory and file exist and are writable by MySQL:

```bash
sudo mkdir -p /var/log/mysql
sudo touch /var/log/mysql/error.log
sudo chown mysql:mysql /var/log/mysql/error.log
```

6. Restart MySQL server:

```bash
sudo systemctl restart mysql
```

Now your MySQL error logs will be written to `/var/log/mysql/error.log`. You can view them with:

```bash
sudo tail -f /var/log/mysql/error.log
```

---


