import paramiko
import time

FG_IP     = "FW IP Address"
FG_USER   = "username with superadmin priv"
FG_PASS   = "password"

POLICY_IDS = [ Policy IDs ]

def send(chan, cmd):
    chan.send(cmd + "\n")
    time.sleep(0.4)

def main():
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    print("Connect To Firewall")
    ssh.connect(FG_IP, username=FG_USER, password=FG_PASS)

    chan = ssh.invoke_shell()
    time.sleep(1)

    # login   
    send(chan, "config vdom")
    send(chan, "edit root")

    print("reset counter")

    for pid in POLICY_IDS:
        print(f"→ Reset Policy {pid}")

        #    reset counter
        send(chan, f"diagnose firewall iprope clear 00100004 {pid}")

    print("\n✅")

    ssh.close()


if __name__ == "__main__":
    main()
