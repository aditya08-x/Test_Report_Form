from createTestReportPackages.api.controller import app


def start_prod_server(host, port, number_of_workers, threads):
    """Start gunicorn production server

    :param host: host IP
    :param port: port number
    :param prod: whether to run production server or not
    :param number_of_workers: number of worker for production server
    :param threads: number of threads for production server
    """
    from gunicorn.app.base import BaseApplication

    class StandaloneApplication(BaseApplication):
        """Wrapper class for base application to provide command line options"""
        def __init__(self, app, options=None):
            self.options = options or {}
            self.application = app
            super().__init__()

        def load_config(self):
            config = {key: value for key, value in self.options.items()
                      if key in self.cfg.settings and value is not None}
            for key, value in config.items():
                self.cfg.set(key.lower(), value)

        def load(self):
            return self.application

    options = {
        'bind': '%s:%s' % (host, port),
        'workers': int(number_of_workers),
        'threads': int(threads)
    }
    StandaloneApplication(app, options).run()


def start_hosting(host, port, prod, number_of_workers, threads):
    """ Start hosting python server

    :param host: host ip
    :param port: port number
    :param prod: whether to run production server or not
    :param number_of_workers: number of worker for production server
    :param threads: number of threads for production server
    """
    try:
        if not prod:
            print("Starting Development Server")
            app.run(debug=False, host=host, port=port)
        else:
            print("Starting Production Server")
            start_prod_server(host=host,
                              port=port,
                              number_of_workers=number_of_workers,
                              threads=threads)
    except Exception as exp_obj:
        raise exp_obj
