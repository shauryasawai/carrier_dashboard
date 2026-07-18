from django.urls import path

from dashboard import views

urlpatterns = [
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path("", views.index, name="index"),
    path("api/process/", views.process_upload, name="process_upload"),
    path("api/bigquery/", views.load_bigquery, name="load_bigquery"),
    path("api/export/", views.export_shipments, name="export_shipments"),
    path("api/ai-summary/", views.ai_summary, name="ai_summary"),
    path("api/sla/", views.sla_config, name="sla_config"),
    path("api/invoices/", views.process_invoices, name="process_invoices"),
    path("api/invoices/awbs/", views.export_invoice_awbs, name="export_invoice_awbs"),
    path("api/master/", views.master_config, name="master_config"),
    path("api/invoices/drive/", views.import_from_drive, name="import_from_drive"),
    path("api/carriers/register/", views.register_carrier, name="register_carrier"),
]
