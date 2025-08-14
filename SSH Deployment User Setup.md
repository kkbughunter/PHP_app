
# SSH Deployment User Setup: `developer1`

## 1. Create the New User
```bash
sudo adduser developer1
````

* Set password when prompted.
* Skip optional user info if not needed.

---

## 2. Configure SSH Access

### Create `.ssh` directory

```bash
sudo mkdir -p /home/developer1/.ssh
sudo chmod 700 /home/developer1/.ssh
sudo chown developer1:developer1 /home/developer1/.ssh
```

### Add Your Public Key

On your **local machine**:

```bash
ssh-keygen -t ed25519 -C "your_email@example.com"
cat ~/.ssh/id_ed25519.pub
```

Copy the output.

On the **server**:

```bash
sudo nano /home/developer1/.ssh/authorized_keys
# Paste the copied public key
sudo chmod 600 /home/developer1/.ssh/authorized_keys
sudo chown developer1:developer1 /home/developer1/.ssh/authorized_keys
```

---

## 3. Test SSH Login

From your **local machine**:

```bash
ssh developer1@<server_ip>
```

* You should connect without being prompted for a password.

---

## 4. (Optional) Restrict SSH to the User

Edit SSH config:

```bash
sudo nano /etc/ssh/sshd_config
```

Add:

```
AllowUsers developer1
```

Restart SSH:

```bash
sudo systemctl restart ssh
```

---

## 5. Deployment Usage

* Use `developer1` for **manual SSH logins** or **automated deployments**.
* For GitHub deployments:

  1. Generate a key as `developer1`:

     ```bash
     ssh-keygen -t ed25519 -C "developer1@server"
     ```
  2. Add `/home/developer1/.ssh/id_ed25519.pub` to your GitHub repo as a **Deploy Key**.
  3. Use `git@github.com:user/repo.git` for pulling code.

---

✅ **Result:**
`developer1` can log in over SSH without a password, and can be safely used for deployment without giving root access.

```
