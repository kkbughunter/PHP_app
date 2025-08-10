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

