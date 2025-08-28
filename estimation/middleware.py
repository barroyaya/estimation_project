# estimation/middleware.py
class ClientMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        client_id = request.session.get('client_id')
        if client_id:
            try:
                from .models import Client
                request.client = Client.objects.get(id=client_id, actif=True)
            except Client.DoesNotExist:
                request.session.flush()
                request.client = None
        else:
            request.client = None

        response = self.get_response(request)
        return response