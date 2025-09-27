from django.contrib import admin
from django.urls import path, include

urlpatterns = [

    path('admin/', admin.site.urls),
    path('api/file/read-data/', include('read_data.urls')),
    path('', include('metrics_data.urls')),
    path('api/ia/', include('intelligent_model.urls'))

]