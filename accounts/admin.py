from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import User

class CustomUserAdmin(UserAdmin):
    list_display = ('username', 'email', 'role', 'department', 'is_staff')
    fieldsets = UserAdmin.fieldsets + (
        ('Role Info', {'fields': ('role', 'department', 'phone', 'profile_pic',
                                    'roll_no', 'year', 'employee_id', 'designation')}),
    )

admin.site.register(User, CustomUserAdmin)