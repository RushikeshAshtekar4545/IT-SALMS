from django.contrib.auth.models import AbstractUser
from django.db import models


class User(AbstractUser):
    class Role(models.TextChoices):
        STUDENT = 'STUDENT', 'Student'
        FACULTY = 'FACULTY', 'Faculty'
        ADMIN = 'ADMIN', 'Admin'

    role = models.CharField(max_length=10, choices=Role.choices)

    department = models.CharField(
        max_length=100,
        default='Information Technology'
    )

    phone = models.CharField(max_length=15, blank=True, null=True)

    profile_pic = models.ImageField(
        upload_to='profile_pics/',
        blank=True,
        null=True
    )

    roll_no = models.CharField(max_length=20, blank=True, null=True)
    year = models.CharField(max_length=10, blank=True, null=True)

    employee_id = models.CharField(
        max_length=20,
        blank=True,
        null=True
    )

    designation = models.CharField(
        max_length=50,
        blank=True,
        null=True
    )

    def __str__(self):
        return f"{self.username} ({self.role})"


class Student(models.Model):
    student_id = models.AutoField(primary_key=True)
    roll_no = models.IntegerField()
    prn = models.CharField(max_length=20, unique=True)
    full_name = models.CharField(max_length=100)
    email = models.EmailField(unique=True)
    mobile = models.CharField(max_length=15, blank=True, null=True)
    gender = models.CharField(max_length=10, blank=True, null=True)
    academic_year = models.CharField(max_length=20, default='2025-26')
    year = models.CharField(max_length=10)
    branch = models.CharField(max_length=20, default='IT')
    division = models.CharField(max_length=1, blank=True, null=True)
    password = models.CharField(max_length=255)
    profile_photo = models.CharField(max_length=255, blank=True, null=True)
    status = models.CharField(max_length=10, default='Active')
    created_at = models.DateTimeField()

    class Meta:
        db_table = 'students'
        managed = False

    def __str__(self):
        return f"{self.prn} - {self.full_name}"


class Faculty(models.Model):
    faculty_id = models.AutoField(primary_key=True)
    employee_id = models.CharField(max_length=20, unique=True)
    full_name = models.CharField(max_length=100)
    order_type = models.CharField(max_length=50, blank=True, null=True)
    designation = models.CharField(max_length=100, blank=True, null=True)
    email = models.EmailField(unique=True)
    mobile = models.CharField(max_length=15, blank=True, null=True)
    password = models.CharField(max_length=255, blank=True, null=True)
    account_created = models.BooleanField(default=False)
    status = models.CharField(max_length=10, default='Active')
    created_at = models.DateTimeField()

    class Meta:
        db_table = 'faculty'
        managed = False

    def __str__(self):
        return f"{self.employee_id} - {self.full_name}"


class Activity(models.Model):
    activity_id = models.AutoField(primary_key=True)
    student = models.ForeignKey(Student, on_delete=models.CASCADE, db_column='student_id')
    activity_type = models.CharField(max_length=20)
    activity_subtype = models.CharField(max_length=100, blank=True, null=True)
    activity_name = models.CharField(max_length=150)
    organizer = models.CharField(max_length=150, blank=True, null=True)
    start_date = models.DateField(blank=True, null=True)
    end_date = models.DateField(blank=True, null=True)
    event_level = models.CharField(max_length=20, blank=True, null=True)
    rank_achieved = models.CharField(max_length=20, blank=True, null=True)
    description = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        db_table = 'activities'
        managed = False

    def __str__(self):
        return self.activity_name


class Certificate(models.Model):
    certificate_id = models.AutoField(primary_key=True)
    student = models.ForeignKey(Student, on_delete=models.CASCADE, db_column='student_id')
    activity = models.ForeignKey(Activity, on_delete=models.CASCADE, db_column='activity_id')
    certificate_file = models.CharField(max_length=255)
    verification_status = models.CharField(max_length=20, default='Pending')
    verified_by = models.IntegerField(blank=True, null=True)
    upload_date = models.DateTimeField(blank=True, null=True)

    class Meta:
        db_table = 'certificates'
        managed = False


class LeaveApplication(models.Model):
    leave_id = models.AutoField(primary_key=True)
    student = models.ForeignKey(Student, on_delete=models.CASCADE, db_column='student_id')
    activity = models.ForeignKey(Activity, on_delete=models.CASCADE, db_column='activity_id')
    reason = models.TextField()
    leave_start = models.DateField()
    leave_end = models.DateField()
    status = models.CharField(max_length=20, default='Pending')
    faculty_remark = models.TextField(blank=True, null=True)
    applied_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        db_table = 'leave_applications'
        managed = False