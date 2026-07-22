from django.conf import settings
settings.configure(USE_X_FORWARDED_HOST=True, ALLOWED_HOSTS=['31.97.227.65'])
from django.http import HttpRequest
r = HttpRequest()
r.META['HTTP_X_FORWARDED_HOST'] = '31.97.227.65'
r.META['HTTP_HOST'] = '31.97.227.65,31.97.227.65'
print(r.get_host())
