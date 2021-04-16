from django.urls import path, re_path
from Handle_Files import views

urlpatterns = [
    path('', views.index),
    path('download/', views.download),
    path('download_pdf', views.download_pdf)
    # path('transform/', views.search),
]
