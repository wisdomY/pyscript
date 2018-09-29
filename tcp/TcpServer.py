from socket import *

def StartServer():
    server = socket(AF_INET, SOCK_STREAM)
    server.bind(("127.0.0.1", 5555))
    server.listen(1)

    sock, addr = server.accept()
    print "[+] connected from ", addr
    sock.send('Hello')

    while True:
        r = sock.recv(1024)
        print "[+] message:", r
        sock.send("you said: " + r)

    server.close()


if __name__ == '__main__':
    # Start a simple server, and loop forever
    StartServer()