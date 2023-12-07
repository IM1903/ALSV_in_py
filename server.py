from http.server import HTTPServer, SimpleHTTPRequestHandler


def serve(port):
    handler = SimpleHTTPRequestHandler
    httpd = HTTPServer(("localhost", port), handler)
    print("解析完成. 访问 http://localhost:{0}".format(port))
    httpd.serve_forever()


if __name__ == "__main__":
    serve(8000)
