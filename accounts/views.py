from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db import connection
from django.http import JsonResponse
from django.conf import settings
from django.utils import timezone
from datetime import date
import os
import re

from PyPDF2 import PdfReader
from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

from .models import User, Student, Faculty, Activity, Certificate, LeaveApplication

def login_view(request):
    if request.method == 'POST':
        role = request.POST.get('role')
        username = request.POST.get('username')
        password = request.POST.get('password')

        user = authenticate(
            request,
            username=username,
            password=password
        )

        if user is not None:
            if user.role == role:
                login(request, user)
                return redirect('dashboard_redirect')
            else:
                return render(request, 'accounts/login.html', {
                    'error': 'Selected role does not match this account.'
                })
        else:
            return render(request, 'accounts/login.html', {
                'error': 'Invalid username or password.'
            })

    return render(request, 'accounts/login.html')


def create_account(request):

    if request.method == 'POST':

        role = request.POST.get('role')
        identifier = request.POST.get('identifier')
        username = request.POST.get('username')
        password = request.POST.get('password')
        confirm_password = request.POST.get('confirm_password')

        if password != confirm_password:
            return render(request, 'accounts/create_account.html', {
                'error': 'Passwords do not match.',
                'role': role,
                'identifier': identifier,
                'username': username,
            })

        if User.objects.filter(username=username).exists():
            return render(request, 'accounts/create_account.html', {
                'error': 'Username already exists.',
                'role': role,
                'identifier': identifier,
                'username': username,
            })

        if role == 'STUDENT':

            student = Student.objects.filter(
                prn=identifier,
                status='Active'
            ).first()

            if student is None:
                return render(request, 'accounts/create_account.html', {
                    'error': 'Student PRN not found or account is inactive.',
                    'role': role,
                    'identifier': identifier,
                })

            if User.objects.filter(
                email=student.email,
                role='STUDENT'
            ).exists():
                return render(request, 'accounts/create_account.html', {
                    'error': 'An account has already been created for this student.',
                    'role': role,
                    'identifier': identifier,
                })

            User.objects.create_user(
                username=username,
                password=password,
                email=student.email,
                role='STUDENT',
                department='Information Technology'
            )

            return render(request, 'accounts/login.html', {
                'success': 'Account created successfully. Please login.'
            })

        elif role == 'FACULTY':

            faculty = Faculty.objects.filter(
                employee_id=identifier,
                status='Active'
            ).first()

            if faculty is None:
                return render(request, 'accounts/create_account.html', {
                    'error': 'Employee ID not found or account is inactive.',
                    'role': role,
                    'identifier': identifier,
                })

            if User.objects.filter(
                employee_id=faculty.employee_id,
                role='FACULTY'
            ).exists():
                return render(request, 'accounts/create_account.html', {
                    'error': 'An account has already been created for this faculty member.',
                    'role': role,
                    'identifier': identifier,
                })

            User.objects.create_user(
                username=username,
                password=password,
                email=faculty.email,
                role='FACULTY',
                department='Information Technology',
                employee_id=faculty.employee_id,
                designation=faculty.designation
            )

            return render(request, 'accounts/login.html', {
                'success': 'Account created successfully. Please login.'
            })

        return render(request, 'accounts/create_account.html', {
            'error': 'Invalid role selected.'
        })

    role = request.GET.get('role', 'STUDENT')

    return render(request, 'accounts/create_account.html', {
        'role': role
    })


def logout_view(request):
    logout(request)
    return redirect('login')


@login_required
def dashboard_redirect(request):
    if request.user.role == 'FACULTY':
        return redirect('faculty_dashboard')

    if request.user.role == 'ADMIN':
        return redirect('admin_dashboard')

    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/student_dashboard.html', {
            'student_name': request.user.username,
            'error': 'No student record linked to this account.'
        })

    activities = Activity.objects.filter(student=student)

    hackathon_count = activities.filter(activity_type='Hackathon').count()
    workshop_count = activities.filter(activity_type='Workshop').count()
    event_count = activities.filter(activity_type='Event').count()
    certificate_count = Certificate.objects.filter(student=student).count()

    leaves = LeaveApplication.objects.filter(student=student)
    pending_count = leaves.filter(status='Pending').count()
    approved_count = leaves.filter(status='Approved').count()
    rejected_count = leaves.filter(status='Rejected').count()

    today = date.today()
    recent_activities = activities.filter(start_date__lte=today).order_by('-start_date')[:3]
    upcoming_activities = activities.filter(start_date__gt=today).order_by('start_date')[:3]
    # Generate initials from student's full name
    name_parts = student.full_name.split()

    initials = ''.join(
    part[0].upper() for part in name_parts[:2]
    )

    context = {
        'student_name': student.full_name,
        'initials': initials,
        'hackathon_count': hackathon_count,
        'workshop_count': workshop_count,
        'event_count': event_count,
        'certificate_count': certificate_count,
        'pending_count': pending_count,
        'approved_count': approved_count,
        'rejected_count': rejected_count,
        'recent_activities': recent_activities,
        'upcoming_activities': upcoming_activities,
    }

    return render(request, 'accounts/student_dashboard.html', context)


@login_required
def apply_leave(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/apply_leave.html', {
            'error': 'No student record linked to this account.'
        })

    if request.method == 'POST':
        activity_type = request.POST.get('activity_type')
        activity_name = request.POST.get('activity_name')
        organizer = request.POST.get('organizer')
        start_date = request.POST.get('start_date')
        end_date = request.POST.get('end_date')
        reason = request.POST.get('reason')

        new_activity = Activity.objects.create(
            student=student,
            activity_type=activity_type,
            activity_name=activity_name,
            organizer=organizer,
            start_date=start_date,
            end_date=end_date,
        )

        LeaveApplication.objects.create(
            student=student,
            activity=new_activity,
            reason=reason,
            leave_start=start_date,
            leave_end=end_date,
            status='Pending',
        )

        return redirect('dashboard_redirect')

    return render(request, 'accounts/apply_leave.html')

@login_required
def my_applications(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/my_applications.html', {
            'error': 'No student record linked to this account.'
        })

    applications = LeaveApplication.objects.filter(student=student).select_related('activity').order_by('-applied_at')

    context = {
        'student_name': student.full_name,
        'applications': applications,
    }

    return render(request, 'accounts/my_applications.html', context)

@login_required
def my_activities(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/my_activities.html', {
            'error': 'No student record linked to this account.'
        })

    activities = Activity.objects.filter(student=student).order_by('-start_date')

    hackathon_count = activities.filter(activity_type='Hackathon').count()
    workshop_count = activities.filter(activity_type='Workshop').count()
    event_count = activities.filter(activity_type='Event').count()

    context = {
        'student_name': student.full_name,
        'activities': activities,
        'hackathon_count': hackathon_count,
        'workshop_count': workshop_count,
        'event_count': event_count,
    }

    return render(request, 'accounts/my_activities.html', context)
    



@login_required
def my_certificates(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/my_certificates.html', {
            'error': 'No student record linked to this account.'
        })

    certificates = Certificate.objects.filter(student=student).select_related('activity').order_by('-upload_date')

    return render(request, 'accounts/my_certificates.html', {
        'certificates': certificates,
    })

@login_required
def my_portfolio(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/my_portfolio.html', {
            'error': 'No student record linked to this account.'
        })

    activities = Activity.objects.filter(student=student)
    certificates = Certificate.objects.filter(student=student)

    hackathon_total = activities.filter(activity_type='Hackathon').count()
    workshop_total = activities.filter(activity_type='Workshop').count()
    event_total = activities.filter(activity_type='Event').count()

    hackathon_certs = certificates.filter(activity__activity_type='Hackathon').count()
    workshop_certs = certificates.filter(activity__activity_type='Workshop').count()
    event_certs = certificates.filter(activity__activity_type='Event').count()

    context = {
        'student': student,
        'hackathon_total': hackathon_total,
        'workshop_total': workshop_total,
        'event_total': event_total,
        'hackathon_certs': hackathon_certs,
        'workshop_certs': workshop_certs,
        'event_certs': event_certs,
        'certificate_total': certificates.count(),
        'grand_total': activities.count(),
    }

    return render(request, 'accounts/my_portfolio.html', context)
@login_required
def profile_view(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/profile.html', {
            'error': 'No student record linked to this account.'
        })

    if request.method == 'POST':
        mobile = request.POST.get('mobile')
        gender = request.POST.get('gender')

        student.mobile = mobile
        student.gender = gender
        student.save()

        return render(request, 'accounts/profile.html', {
            'student': student,
            'success': 'Profile updated successfully.'
        })

    return render(request, 'accounts/profile.html', {
        'student': student,
    })
    
@login_required
def faculty_dashboard(request):
    faculty = Faculty.objects.filter(email=request.user.email).first()

    if faculty is None:
        return render(request, 'accounts/faculty_dashboard.html', {
            'error': 'No faculty record linked to this account.'
        })

    pending_leaves = LeaveApplication.objects.filter(status='Pending').select_related('student', 'activity').order_by('-applied_at')
    pending_certificates = Certificate.objects.filter(verification_status='Pending').select_related('student', 'activity').order_by('-upload_date')

    total_students = Student.objects.filter(status='Active').count()

    from datetime import date
    today = date.today()
    activities_this_month = Activity.objects.filter(
        start_date__year=today.year,
        start_date__month=today.month
    ).count()

    context = {
        'faculty_name': faculty.full_name,
        'pending_leaves_count': pending_leaves.count(),
        'pending_certificates_count': pending_certificates.count(),
        'total_students': total_students,
        'activities_this_month': activities_this_month,
        'pending_leaves': pending_leaves[:4],
        'pending_certificates': pending_certificates[:4],
    }

    return render(request, 'accounts/faculty_dashboard.html', context)

@login_required
def leave_approval(request):
    faculty = Faculty.objects.filter(email=request.user.email).first()

    if faculty is None:
        return render(request, 'accounts/leave_approval.html', {
            'error': 'No faculty record linked to this account.'
        })

    if request.method == 'POST':
        leave_id = request.POST.get('leave_id')
        action = request.POST.get('action')
        remark = request.POST.get('remark', '')

        leave = LeaveApplication.objects.filter(leave_id=leave_id).first()
        if leave:
            leave.status = 'Approved' if action == 'approve' else 'Rejected'
            leave.faculty_remark = remark
            leave.save()

        return redirect('leave_approval')

    pending_leaves = LeaveApplication.objects.filter(status='Pending').select_related('student', 'activity').order_by('-applied_at')

    return render(request, 'accounts/leave_approval.html', {
        'pending_leaves': pending_leaves,
    })

@login_required
def certificate_verification(request):
    faculty = Faculty.objects.filter(email=request.user.email).first()

    if faculty is None:
        return render(request, 'accounts/certificate_verification.html', {
            'error': 'No faculty record linked to this account.'
        })

    if request.method == 'POST':
        certificate_id = request.POST.get('certificate_id')
        action = request.POST.get('action')

        certificate = Certificate.objects.filter(certificate_id=certificate_id).first()
        if certificate:
            certificate.verification_status = 'Verified' if action == 'verify' else 'Rejected'
            certificate.verified_by = faculty.faculty_id
            certificate.save()

        return redirect('certificate_verification')

    pending_certificates = Certificate.objects.filter(verification_status='Pending').select_related('student', 'activity').order_by('-upload_date')

    return render(request, 'accounts/certificate_verification.html', {
        'pending_certificates': pending_certificates,
    })
    
@login_required
def faculty_profile(request):
    faculty = Faculty.objects.filter(email=request.user.email).first()

    if faculty is None:
        return render(request, 'accounts/faculty_profile.html', {
            'error': 'No faculty record linked to this account.'
        })

    if request.method == 'POST':
        mobile = request.POST.get('mobile')

        faculty.mobile = mobile
        faculty.save()

        return render(request, 'accounts/faculty_profile.html', {
            'faculty': faculty,
            'success': 'Profile updated successfully.'
        })

    return render(request, 'accounts/faculty_profile.html', {
        'faculty': faculty,
    })

@login_required
def admin_dashboard(request):
    if request.user.role != 'ADMIN':
        return redirect('dashboard_redirect')

    total_students = Student.objects.filter(status='Active').count()
    total_faculty = Faculty.objects.filter(status='Active').count()
    total_activities = Activity.objects.count()
    total_certificates = Certificate.objects.count()

    hackathon_count = Activity.objects.filter(activity_type='Hackathon').count()
    workshop_count = Activity.objects.filter(activity_type='Workshop').count()
    event_count = Activity.objects.filter(activity_type='Event').count()

    recent_activities = Activity.objects.select_related('student').order_by('-created_at')[:5]

    context = {
        'total_students': total_students,
        'total_faculty': total_faculty,
        'total_activities': total_activities,
        'total_certificates': total_certificates,
        'hackathon_count': hackathon_count,
        'workshop_count': workshop_count,
        'event_count': event_count,
        'recent_activities': recent_activities,
    }

    return render(request, 'accounts/admin_dashboard.html', context)

def manage_students(request):
    students = Student.objects.all().order_by('roll_no')

    return render(
        request,
        'accounts/manage_students.html',
        {
            'students': students
        }
    )


def manage_faculty(request):
    faculty = Faculty.objects.all().order_by('employee_id')

    return render(
        request,
        'accounts/manage_faculty.html',
        {
            'faculty': faculty
        }
    )

def add_faculty(request):

    if request.method == 'POST':

        employee_id = request.POST.get('employee_id')
        full_name = request.POST.get('full_name')
        designation = request.POST.get('designation')
        email = request.POST.get('email')
        mobile = request.POST.get('mobile')
        status = request.POST.get('status') or 'Active'

        # Check Employee ID
        if Faculty.objects.filter(employee_id=employee_id).exists():
            return render(request, 'accounts/add_faculty.html', {
                'error': 'Employee ID already exists.',
                'employee_id': employee_id,
                'full_name': full_name,
                'designation': designation,
                'email': email,
                'mobile': mobile,
                'status': status,
            })

        # Check Email
        if Faculty.objects.filter(email=email).exists():
            return render(request, 'accounts/add_faculty.html', {
                'error': 'Email already exists.',
                'employee_id': employee_id,
                'full_name': full_name,
                'designation': designation,
                'email': email,
                'mobile': mobile,
                'status': status,
            })

        Faculty.objects.create(
            employee_id=employee_id,
            full_name=full_name,
            designation=designation,
            email=email,
            mobile=mobile,
            status=status,
            account_created=False,
        )

        return redirect('manage_faculty')

    return render(request, 'accounts/add_faculty.html')

def view_faculty(request, faculty_id):

    faculty = Faculty.objects.get(faculty_id=faculty_id)

    return render(
        request,
        'accounts/view_faculty.html',
        {
            'faculty': faculty
        }
    )

def edit_faculty(request, faculty_id):
    faculty = Faculty.objects.get(faculty_id=faculty_id)

    if request.method == 'POST':
        faculty.full_name = request.POST.get('full_name')
        faculty.employee_id = request.POST.get('employee_id')
        faculty.designation = request.POST.get('designation')
        faculty.email = request.POST.get('email')
        faculty.mobile = request.POST.get('mobile')
        faculty.status = request.POST.get('status')

        faculty.save()

        return redirect('manage_faculty')

    return render(
        request,
        'accounts/edit_faculty.html',
        {
            'faculty': faculty
        }
    )

@login_required
def upload_certificate(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return render(request, 'accounts/upload_certificate.html', {
            'error': 'No student record linked to this account.'
        })

    student_activities = Activity.objects.filter(student=student).order_by('-start_date')

    if request.method == 'POST':
        activity_id = request.POST.get('activity_id')
        uploaded_file = request.FILES.get('certificate_file')

        activity = Activity.objects.filter(activity_id=activity_id, student=student).first()

        if activity is None or uploaded_file is None:
            return render(request, 'accounts/upload_certificate.html', {
                'student_activities': student_activities,
                'error': 'Please select an activity and choose a file.'
            })

        # ---------- SAVE FILE TEMPORARILY SO WE CAN EXTRACT + LATER CONFIRM ----------
        temp_dir = os.path.join(settings.MEDIA_ROOT, 'temp_certificates')
        os.makedirs(temp_dir, exist_ok=True)

        temp_path = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_path, 'wb+') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        # ---------- EXTRACTION ----------
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        extracted_text = ""

        try:
            if file_extension == ".pdf":
                reader = PdfReader(temp_path)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        extracted_text += page_text + "\n"

            elif file_extension in [".jpg", ".jpeg", ".png"]:
                image = Image.open(temp_path)
                extracted_text = pytesseract.image_to_string(image)

            else:
                os.remove(temp_path)
                return render(request, 'accounts/upload_certificate.html', {
                    'student_activities': student_activities,
                    'error': 'Only PDF, JPG, JPEG and PNG files are allowed.'
                })
        except Exception as e:
            os.remove(temp_path)
            return render(request, 'accounts/upload_certificate.html', {
                'student_activities': student_activities,
                'error': f'Unable to read certificate: {str(e)}'
            })

        text = re.sub(r"\s+", " ", extracted_text).strip()

        prn_match = re.search(
            r"\bPRN\s*[:\-]?\s*([A-Za-z0-9\-]{5,20})\b", text, re.IGNORECASE
        )
        extracted_prn = prn_match.group(1).strip() if prn_match else ""

        name_match = re.search(
            r"(?:Participant\s*Name|Student\s*Name|Candidate\s*Name|Name)"
            r"\s*[:\-]?\s*([A-Za-z][A-Za-z .]{1,80})"
            r"(?=\s+(?:PRN|Roll|Branch|Year|Class|Academic|Department)\b|$)",
            text, re.IGNORECASE
        )
        extracted_name = name_match.group(1).strip(" :-") if name_match else ""

        branch_match = re.search(
            r"(?:Branch|Department)"
            r"\s*[:\-]?\s*([A-Za-z][A-Za-z &./-]{1,50})"
            r"(?=\s+(?:Year|Class|Academic|PRN|Roll)\b|$)",
            text, re.IGNORECASE
        )
        extracted_branch = branch_match.group(1).strip(" :-") if branch_match else ""

        year_match = re.search(
            r"(?:Year|Class)"
            r"\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9 ._-]{0,30})"
            r"(?=\s+(?:Branch|Department|Academic|PRN|Roll)\b|$)",
            text, re.IGNORECASE
        )
        extracted_year = year_match.group(1).strip(" :-") if year_match else ""

        # ---------- PRN CHECK: CANCEL UPLOAD IF MISMATCH ----------
        if extracted_prn and str(extracted_prn).strip() != str(student.prn).strip():
            os.remove(temp_path)
            return render(request, 'accounts/upload_certificate.html', {
                'student_activities': student_activities,
                'error': f'Upload cancelled: The PRN on the certificate ({extracted_prn}) does not match your account ({student.prn}).'
            })

        # ---------- PRN OK (OR NOT FOUND) -> SHOW PREVIEW + CONFIRM BUTTON ----------
        request.session['pending_certificate'] = {
            'temp_filename': uploaded_file.name,
            'activity_id': activity_id,
        }

        return render(request, 'accounts/upload_certificate.html', {
            'student_activities': student_activities,
            'extracted_name': extracted_name or 'Not found',
            'extracted_prn': extracted_prn or 'Not found',
            'extracted_branch': extracted_branch or 'Not found',
            'extracted_year': extracted_year or 'Not found',
            'show_confirm': True,
        })

    return render(request, 'accounts/upload_certificate.html', {
        'student_activities': student_activities,
    })


@login_required
def confirm_certificate(request):
    student = Student.objects.filter(email=request.user.email).first()

    if student is None or request.method != 'POST':
        return redirect('upload_certificate')

    pending = request.session.get('pending_certificate')

    if not pending:
        return redirect('upload_certificate')

    temp_path = os.path.join(settings.MEDIA_ROOT, 'temp_certificates', pending['temp_filename'])

    if not os.path.exists(temp_path):
        del request.session['pending_certificate']
        return render(request, 'accounts/upload_certificate.html', {
            'student_activities': Activity.objects.filter(student=student).order_by('-start_date'),
            'error': 'Certificate file expired, please upload again.'
        })

    activity = Activity.objects.filter(activity_id=pending['activity_id'], student=student).first()

    if activity is None:
        os.remove(temp_path)
        del request.session['pending_certificate']
        return redirect('upload_certificate')

    save_dir = os.path.join(settings.MEDIA_ROOT, 'certificates')
    os.makedirs(save_dir, exist_ok=True)

    final_path = os.path.join(save_dir, pending['temp_filename'])
    os.replace(temp_path, final_path)

    Certificate.objects.create(
        student=student,
        activity=activity,
        certificate_file=f'certificates/{pending["temp_filename"]}',
        verification_status='Pending',
        upload_date=timezone.now(),
    )

    del request.session['pending_certificate']

    return redirect('my_certificates')


@login_required
def cancel_certificate(request):
    pending = request.session.pop('pending_certificate', None)

    if pending:
        temp_path = os.path.join(settings.MEDIA_ROOT, 'temp_certificates', pending['temp_filename'])
        if os.path.exists(temp_path):
            os.remove(temp_path)

    return redirect('upload_certificate')

@login_required
def extract_certificate(request):

    if request.method != "POST":
        return JsonResponse({
            "success": False,
            "error": "Invalid request method."
        })

    student = Student.objects.filter(email=request.user.email).first()

    if student is None:
        return JsonResponse({
            "success": False,
            "error": "No student record linked to this account."
        })

    uploaded_file = request.FILES.get("certificate_file")

    if uploaded_file is None:
        return JsonResponse({
            "success": False,
            "error": "Please select a certificate."
        })

    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    extracted_text = ""

    try:
        if file_extension == ".pdf":
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    extracted_text += page_text + "\n"

        elif file_extension in [".jpg", ".jpeg", ".png"]:
            image = Image.open(uploaded_file)
            extracted_text = pytesseract.image_to_string(image)

        else:
            return JsonResponse({
                "success": False,
                "error": "Only PDF, JPG, JPEG and PNG files are allowed."
            })

    except Exception as e:
        return JsonResponse({
            "success": False,
            "error": f"Unable to read certificate: {str(e)}"
        })

    text = re.sub(r"\s+", " ", extracted_text).strip()

    if not text:
        return JsonResponse({
            "success": False,
            "error": "No readable text was found in the certificate."
        })

    prn_match = re.search(
        r"\bPRN\s*[:\-]?\s*([A-Za-z0-9\-]{5,20})\b",
        text,
        re.IGNORECASE
    )
    extracted_prn = prn_match.group(1).strip() if prn_match else ""

    name_match = re.search(
        r"(?:Participant\s*Name|Student\s*Name|Candidate\s*Name|Name)"
        r"\s*[:\-]?\s*([A-Za-z][A-Za-z .]{1,80})"
        r"(?=\s+(?:PRN|Roll|Branch|Year|Class|Academic|Department)\b|$)",
        text,
        re.IGNORECASE
    )
    extracted_name = name_match.group(1).strip(" :-") if name_match else ""

    branch_match = re.search(
        r"(?:Branch|Department)"
        r"\s*[:\-]?\s*([A-Za-z][A-Za-z &./-]{1,50})"
        r"(?=\s+(?:Year|Class|Academic|PRN|Roll)\b|$)",
        text,
        re.IGNORECASE
    )
    extracted_branch = branch_match.group(1).strip(" :-") if branch_match else ""

    year_match = re.search(
        r"(?:Year|Class)"
        r"\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9 ._-]{0,30})"
        r"(?=\s+(?:Branch|Department|Academic|PRN|Roll)\b|$)",
        text,
        re.IGNORECASE
    )
    extracted_year = year_match.group(1).strip(" :-") if year_match else ""

    academic_year_match = re.search(
        r"(?:Academic\s*Year|Academic\s*Session)"
        r"\s*[:\-]?\s*([0-9]{4}\s*[-/]\s*[0-9]{2,4})",
        text,
        re.IGNORECASE
    )
    extracted_academic_year = academic_year_match.group(1).strip() if academic_year_match else ""

    if extracted_prn:
        if str(extracted_prn).strip() != str(student.prn).strip():
            return JsonResponse({
                "success": False,
                "error": "The PRN found on the certificate does not match your account."
            })

    return JsonResponse({
        "success": True,
        "participant_name": extracted_name,
        "prn": extracted_prn,
        "academic_year": extracted_academic_year,
        "year": extracted_year,
        "branch": extracted_branch,
    })