from django.urls import path

from dashboard import views

urlpatterns = [
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("", views.index, name="index"),
    path("api/process/", views.process_upload, name="process_upload"),
    path("api/bigquery/", views.load_bigquery, name="load_bigquery"),
]
