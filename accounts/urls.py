from django.urls import path
from . import views

urlpatterns = [
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('dashboard/', views.dashboard_redirect, name='dashboard_redirect'),
    path('apply-leave/', views.apply_leave, name='apply_leave'),
    path('my-applications/', views.my_applications, name='my_applications'),
    path('my-activities/', views.my_activities, name='my_activities'),
    path('upload-certificate/', views.upload_certificate, name='upload_certificate'),
    path('my-certificates/', views.my_certificates, name='my_certificates'),
    path('my-portfolio/', views.my_portfolio, name='my_portfolio'),
    path('profile/', views.profile_view, name='profile'),
    path('faculty-dashboard/', views.faculty_dashboard, name='faculty_dashboard'),
    path('leave-approval/', views.leave_approval, name='leave_approval'),
    path('create-account/', views.create_account, name='create_account'),
    path('certificate-verification/', views.certificate_verification, name='certificate_verification'),
    path('faculty-profile/', views.faculty_profile, name='faculty_profile'),
    path('admin-dashboard/', views.admin_dashboard, name='admin_dashboard'),
    path('manage-students/', views.manage_students, name='manage_students'),
    path('manage-faculty/', views.manage_faculty, name='manage_faculty'),
    path('add-faculty/', views.add_faculty, name='add_faculty'),
    path('view-faculty/<int:faculty_id>/', views.view_faculty, name='view_faculty'),
    path('edit-faculty/<int:faculty_id>/', views.edit_faculty, name='edit_faculty'),
    path('extract-certificate/', views.extract_certificate, name='extract_certificate'),
    path('confirm-certificate/', views.confirm_certificate, name='confirm_certificate'),
    path('cancel-certificate/', views.cancel_certificate, name='cancel_certificate'),
]