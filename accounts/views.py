from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db import connection
from django.db import transaction
from django.http import JsonResponse, FileResponse
from django.conf import settings
from django.utils import timezone
from datetime import date
from datetime import datetime
from django.http import HttpResponse
from openpyxl import Workbook
import io
import json
import time
import fitz
from openpyxl import Workbook, load_workbook
import os
import re


from PyPDF2 import PdfReader
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import pytesseract

# Use the local Windows Tesseract installation when it exists.
# On Render/Linux, pytesseract will use the system Tesseract command.
WINDOWS_TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if os.path.exists(WINDOWS_TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = WINDOWS_TESSERACT_PATH
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

# ============================================================
# CERTIFICATE EXCEL SETUP
# ============================================================

CERTIFICATE_EXCEL_HEADERS = [
    "Year",
    "Date",
    "Time",
    "PRN",
    "Participant Name",
    "Branch",
    "Event Type",
    "College Name",
]


def append_certificate_to_excel(data):
    """Append one certificate record to the common Excel file."""

    excel_directory = os.path.join(
        settings.MEDIA_ROOT,
        "excel"
    )

    os.makedirs(
        excel_directory,
        exist_ok=True
    )

    excel_path = os.path.join(
        excel_directory,
        "IT_SALMS_Certificates.xlsx"
    )

    # Open existing workbook or create a new one
    if os.path.exists(excel_path):
        workbook = load_workbook(excel_path)
        worksheet = workbook.active
    else:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Certificates"

        worksheet.append(CERTIFICATE_EXCEL_HEADERS)

    # Add certificate data
    worksheet.append([
        data.get("year", ""),
        data.get("date", ""),
        data.get("time", ""),
        data.get("prn", ""),
        data.get("participant_name", ""),
        data.get("branch", ""),
        data.get("event_type", ""),
        data.get("college_name", ""),
    ])

    # Adjust column widths
    for column_cells in worksheet.columns:

        maximum_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            if cell.value is not None:
                maximum_length = max(
                    maximum_length,
                    len(str(cell.value))
                )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            maximum_length + 2,
            40
        )

    workbook.save(excel_path)

    return excel_path

# ============================================================
# CERTIFICATE TEXT EXTRACTION
# ============================================================

def clean_certificate_text(text):
    """Clean PDF/OCR text while preserving useful line breaks."""

    if not text:
        return ""

    text = text.replace("\r", "\n")
    text = text.replace("\x00", " ")

    cleaned_lines = []

    for line in text.split("\n"):
        line = re.sub(r"\s+", " ", line).strip()

        if line:
            cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


def normalized_certificate_text(text):
    """Convert certificate text into one searchable line."""

    return re.sub(
        r"\s+",
        " ",
        clean_certificate_text(text)
    ).strip()


def _prepare_ocr_images(image):
    """
    Create multiple OCR-friendly versions of a certificate image.

    This is especially useful for:
    - WhatsApp compressed images
    - screenshots
    - low-resolution certificates
    - dark/light backgrounds
    - scanned certificates
    """

    images = []

    try:
        image = image.convert("RGB")

        # ----------------------------------------------------
        # ORIGINAL
        # ----------------------------------------------------
        images.append(image)

        # ----------------------------------------------------
        # UPSCALE
        # ----------------------------------------------------
        width, height = image.size

        if width < 1800:
            scale = 2
        else:
            scale = 1

        if scale > 1:
            upscaled = image.resize(
                (width * scale, height * scale),
                Image.Resampling.LANCZOS
            )
        else:
            upscaled = image

        images.append(upscaled)

        # ----------------------------------------------------
        # GRAYSCALE + CONTRAST
        # ----------------------------------------------------
        gray = ImageOps.grayscale(upscaled)

        gray = ImageOps.autocontrast(gray)

        contrast = ImageEnhance.Contrast(gray).enhance(2.0)

        sharp = contrast.filter(
            ImageFilter.SHARPEN
        )

        images.append(sharp)

        # ----------------------------------------------------
        # THRESHOLD VERSION
        # ----------------------------------------------------
        threshold = sharp.point(
            lambda pixel: 255 if pixel > 160 else 0
        )

        images.append(threshold)

        # ----------------------------------------------------
        # SOFT THRESHOLD
        # ----------------------------------------------------
        soft_threshold = sharp.point(
            lambda pixel: 255 if pixel > 190 else 0
        )

        images.append(soft_threshold)

    except Exception as error:
        print("OCR image preparation error:", error)

        images = [image]

    return images


def _ocr_image_variants(image):
    """
    Run Tesseract using several certificate-friendly configurations.
    """

    results = []

    prepared_images = _prepare_ocr_images(image)

    psm_modes = [
        3,   # Automatic page segmentation
        6,   # Uniform block of text
        11,  # Sparse text
        12,  # Sparse text with OSD
    ]

    for prepared_image in prepared_images:

        for psm in psm_modes:

            try:

                text = pytesseract.image_to_string(
                    prepared_image,
                    config=f"--oem 3 --psm {psm}"
                )

                text = clean_certificate_text(text)

                if text:
                    results.append(text)

            except Exception as error:

                print(
                    f"OCR failed for PSM {psm}:",
                    error
                )

    return results


def _score_ocr_text(text):
    """
    Give an OCR result a score based on useful certificate information.

    A result containing several certificate fields is preferred over
    an OCR result containing mostly random WhatsApp-image noise.
    """

    if not text:
        return 0

    score = 0

    normalized = re.sub(
        r"\s+",
        " ",
        text
    ).lower()

    useful_words = [
        "certificate",
        "certify",
        "participant",
        "student",
        "candidate",
        "name",
        "prn",
        "registration",
        "enrollment",
        "branch",
        "department",
        "date",
        "time",
        "year",
        "event",
        "activity",
        "college",
        "institute",
        "university",
        "workshop",
        "hackathon",
        "seminar",
        "webinar",
        "competition",
        "training",
        "internship",
        "organized",
        "participated",
    ]

    for word in useful_words:

        if word in normalized:
            score += 3

    # Date-like information
    if re.search(
        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
        normalized
    ):
        score += 8

    if re.search(
        r"\b20\d{2}\b",
        normalized
    ):
        score += 5

    # PRN / registration-like number
    if re.search(
        r"\b\d{6,10}\b",
        normalized
    ):
        score += 5

    # A reasonable amount of text is useful.
    score += min(
        len(normalized) // 100,
        10
    )

    return score


def _select_best_ocr_text(results):
    """
    Select the OCR result containing the most useful certificate
    information.
    """

    if not results:
        return ""

    unique_results = []

    seen = set()

    for result in results:

        cleaned = clean_certificate_text(result)

        if not cleaned:
            continue

        key = re.sub(
            r"\s+",
            " ",
            cleaned
        ).lower()

        if key not in seen:

            seen.add(key)
            unique_results.append(cleaned)

    if not unique_results:
        return ""

    unique_results.sort(
        key=_score_ocr_text,
        reverse=True
    )

    return unique_results[0]


def extract_certificate_text(file_path):
    """
    Extract text from PDF or image.

    Supports:
    - normal PDFs
    - scanned PDFs
    - JPG/JPEG
    - PNG
    - WhatsApp-compressed images
    - screenshots
    """

    extension = os.path.splitext(
        file_path
    )[1].lower()

    # ========================================================
    # PDF
    # ========================================================

    if extension == ".pdf":

        native_text_parts = []

        try:

            reader = PdfReader(file_path)

            for page in reader.pages:

                page_text = page.extract_text() or ""

                if page_text.strip():

                    native_text_parts.append(
                        page_text
                    )

        except Exception as error:

            print(
                "PDF native extraction error:",
                error
            )

        native_text = clean_certificate_text(
            "\n".join(native_text_parts)
        )

        # If the PDF already contains good text, still use OCR
        # when the text appears too short or incomplete.
        native_score = _score_ocr_text(
            native_text
        )

        # ====================================================
        # OCR scanned PDF
        # ====================================================

        ocr_results = []

        try:

            pdf_document = fitz.open(
                file_path
            )

            for page in pdf_document:

                # Render at high resolution.
                matrix = fitz.Matrix(
                    2.5,
                    2.5
                )

                pix = page.get_pixmap(
                    matrix=matrix,
                    alpha=False
                )

                image_bytes = pix.tobytes(
                    "png"
                )

                image = Image.open(
                    io.BytesIO(image_bytes)
                )

                page_results = _ocr_image_variants(
                    image
                )

                ocr_results.extend(
                    page_results
                )

            pdf_document.close()

        except Exception as error:

            print(
                "PDF OCR error:",
                error
            )

        best_ocr_text = _select_best_ocr_text(
            ocr_results
        )

        ocr_score = _score_ocr_text(
            best_ocr_text
        )

        # Prefer whichever contains more useful certificate
        # information.
        if ocr_score > native_score:
            return best_ocr_text

        if native_text:
            return native_text

        return best_ocr_text

    # ========================================================
    # IMAGE
    # ========================================================

    if extension in [
        ".jpg",
        ".jpeg",
        ".png"
    ]:

        try:

            image = Image.open(
                file_path
            )

            results = _ocr_image_variants(
                image
            )

            return _select_best_ocr_text(
                results
            )

        except Exception as error:

            print(
                "Image OCR error:",
                error
            )

            return ""

    return ""

def _clean_ocr_value(value):
    """Clean and normalize OCR extracted text."""

    if not value:
        return ""

    value = re.sub(r"\s+", " ", value)

    value = value.strip(
        " \t\r\n:,-"
    )

    return value.strip()

def extract_prn(text):
    """Extract PRN / registration / enrollment number."""
    if not text:
        return ""

    patterns = [
        r"\bPRN\s*(?:NO\.?|NUMBER|ID)?\s*[:#\-]?\s*([A-Z0-9]{5,15})\b",
        r"\b(?:Registration|Regn?|Enrollment|Enrolment)\s*(?:No\.?|Number|ID)?\s*[:#\-]?\s*([A-Z0-9]{5,15})\b",
        r"\bRoll\s*(?:No\.?|Number|ID)?\s*[:#\-]?\s*([A-Z0-9]{5,15})\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = _clean_ocr_value(match.group(1))
            if value:
                return value

    return ""


def extract_participant_name(text):
    """Extract participant/student/candidate name from certificate text."""

    if not text:
        return ""

    candidates = []

    # Normalize OCR whitespace first.
    normalized = re.sub(r"\s+", " ", text).strip()

    # =========================================================
    # 1. "This is to certify that, NAME has successfully..."
    # =========================================================
    patterns = [

        r"This\s+is\s+to\s+certif\w*\s+that"
        r"\s*[,:\-]?\s*"
        r"[_\-|]*\s*"
        r"(?:Mr\.?/Ms\.?|Mr\.?|Ms\.?|Mrs\.?|Miss)?\s*"
        r"([A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})"
        r"\s+(?=[a-z]+\s+)?"
        r"(?:has|is|was|have)\b",

        # More tolerant version for OCR junk after the name
        r"This\s+is\s+to\s+certif\w*\s+that"
        r"\s*[,:\-]?\s*"
        r"[_\-|]*\s*"
        r"([A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})"
        r"\s+[a-z]{1,3}\s+"
        r"(?:has|is|was|have)\b",

        # Explicit labels
        r"(?:Participant\s*Name|Student\s*Name|Candidate\s*Name)"
        r"\s*[:\-]?\s*"
        r"(?:Mr\.?/Ms\.?|Mr\.?|Ms\.?|Mrs\.?|Miss)?\s*"
        r"([A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})",

        # Name: ...
        r"\bName\s*[:\-]\s*"
        r"(?:Mr\.?/Ms\.?|Mr\.?|Ms\.?|Mrs\.?|Miss)?\s*"
        r"([A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})",

        # awarded/presented/given/issued to
        r"(?:awarded|presented|given|issued)\s+to\s+"
        r"(?:Mr\.?/Ms\.?|Mr\.?|Ms\.?|Mrs\.?|Miss)?\s*"
        r"([A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})",
    ]

    for pattern in patterns:

        for match in re.finditer(
            pattern,
            normalized,
            re.IGNORECASE
        ):

            name = _clean_ocr_value(match.group(1))

            # Remove leading OCR symbols
            name = re.sub(
                r"^[^A-Za-z]+",
                "",
                name
            )

            # Remove titles
            name = re.sub(
                r"^(?:Mr\.?/Ms\.?|Mr|Ms|Mrs|Miss)\.?\s+",
                "",
                name,
                flags=re.IGNORECASE
            )

            # Remove obvious OCR garbage from the end
            name = re.sub(
                r"\s+(?:r|a|ra|the|of)$",
                "",
                name,
                flags=re.IGNORECASE
            )

            name = _clean_ocr_value(name)

            if name:
                candidates.append(name)

    # =========================================================
    # 2. Direct line-based extraction
    #
    # Handles:
    #
    # This is to certily that,
    # amane Anushka Ramesh
    # has successfully participated
    # =========================================================
    lines = text.splitlines()

    for index, line in enumerate(lines):

        clean_line = _clean_ocr_value(line)

        if not clean_line:
            continue

        if re.search(
            r"This\s+is\s+to\s+certif",
            clean_line,
            re.IGNORECASE
        ):

            # Search the next few lines for the participant name.
            for next_index in range(
                index + 1,
                min(index + 4, len(lines))
            ):

                possible_name = _clean_ocr_value(
                    lines[next_index]
                )

                if not possible_name:
                    continue

                # Stop if we reach the sentence itself.
                if re.search(
                    r"has\s+successfully|has\s+participated",
                    possible_name,
                    re.IGNORECASE
                ):
                    break

                # Remove OCR symbols.
                possible_name = re.sub(
                    r"^[^A-Za-z]+",
                    "",
                    possible_name
                )

                possible_name = re.sub(
                    r"[^A-Za-z.'\-\s]",
                    "",
                    possible_name
                )

                words = possible_name.split()

                # A normal person's name generally has at least
                # two words.
                if len(words) >= 2 and len(words) <= 6:

                    # Reject obvious certificate text.
                    blocked = {
                        "certificate",
                        "participation",
                        "hackathon",
                        "department",
                        "college",
                        "institute",
                        "technology",
                        "organized",
                        "sponsored",
                    }

                    if not any(
                        word.lower() in blocked
                        for word in words
                    ):
                        candidates.append(
                            " ".join(words)
                        )

    # =========================================================
    # 3. Remove duplicates
    # =========================================================
    unique_candidates = []

    for candidate in candidates:

        candidate = _clean_ocr_value(candidate)

        if not candidate:
            continue

        if candidate.lower() not in [
            item.lower()
            for item in unique_candidates
        ]:
            unique_candidates.append(candidate)

    if not unique_candidates:
        return ""

    # =========================================================
    # 4. Prefer the most complete candidate
    # =========================================================
    unique_candidates.sort(
        key=lambda value: (
            len(value.split()),
            len(value)
        ),
        reverse=True
    )

    return unique_candidates[0]

def _normalize_branch(branch):
    """Normalize common branch names and OCR variations."""

    if not branch:
        return ""

    branch = _clean_ocr_value(branch)

    # Common OCR error: AI -> Al
    branch = re.sub(
        r"\bCSE\s*\(\s*Al\s*&\s*ML\s*\)",
        "CSE (AI & ML)",
        branch,
        flags=re.IGNORECASE
    )

    # Normalize common IT branch variations
    branch = re.sub(
        r"\bInformation\s+Technology\b",
        "Information Technology",
        branch,
        flags=re.IGNORECASE
    )

    # Normalize CSE
    branch = re.sub(
        r"\bComputer\s+Science\s+and\s+Engineering\b",
        "Computer Science and Engineering",
        branch,
        flags=re.IGNORECASE
    )

    return branch.strip()

def extract_branch(text):
    """Extract branch / department / stream."""
    if not text:
        return ""

    patterns = [
        r"\b(?:Branch|Branch\s*Name|Stream)\s*[:\-]?\s*"
        r"([A-Za-z][A-Za-z0-9 &()./'-]{1,80})"
        r"(?=\s+(?:Year|Class|Date|Time|PRN|Roll|Name|Participant|Event|College|Institute|University)\b|$)",

        r"\b(?:Department|Dept\.?)\s*[:\-]\s*"
        r"([A-Za-z][A-Za-z0-9 &()./'-]{1,80})"
        r"(?=\s+(?:Year|Class|Date|Time|PRN|Roll|Name|Participant|Event|College|Institute|University)\b|$)",

        r"\bDepartment\s+of\s+"
        r"([A-Za-z][A-Za-z0-9 &()./'-]{1,80}?)"
        r"(?=\s+(?:has|is|was|from|on|at|in|and)\b|[,.;\n]|$)",

        r"\bDept\.?\s+of\s+"
        r"([A-Za-z][A-Za-z0-9 &()./'-]{1,80}?)"
        r"(?=\s+(?:has|is|was|from|on|at|in|and)\b|[,.;\n]|$)",

        r"\b(Information\s+Technology)\b",
        r"\b(Computer\s+Science\s+and\s+Engineering)\b",
        r"\b(CSE\s*\([^)]*\))\b",
        r"\b(CSE)\b",
        r"\b(IT)\b",
        r"\b(Electronics\s+and\s+Telecommunication)\b",
        r"\b(Electronics\s+and\s+Communication)\b",
        r"\b(ENTC)\b",
        r"\b(Mechanical\s+Engineering)\b",
        r"\b(Civil\s+Engineering)\b",
        r"\b(Electrical\s+Engineering)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            branch = _normalize_branch(match.group(1))

            branch = re.split(
                r"\s+(?:of|has|is|was|from|on|at|in|and)\b",
                branch,
                maxsplit=1,
                flags=re.IGNORECASE
            )[0]

            branch = _clean_ocr_value(branch)

            if branch:
                return branch

    return ""


def extract_year(text):
    """Extract academic year or explicit certificate/event year."""

    if not text:
        return ""

    patterns = [

        # Academic Year / Class
        r"(?:Academic\s+)?(?:Year|Class)\s*[:\-]?\s*"
        r"(1st\s*Year|2nd\s*Year|3rd\s*Year|4th\s*Year|"
        r"First\s*Year|Second\s*Year|Third\s*Year|Fourth\s*Year)",

        # of Third year
        r"\b(of\s+)?"
        r"(1st\s*year|2nd\s*year|3rd\s*year|4th\s*year)\b",

        r"\b(of\s+)?"
        r"(first\s+year|second\s+year|third\s+year|fourth\s+year)\b",

        # FY / SY / TY
        r"\b(FY|SY|TY|Final\s*Year)\b",

        # B.Tech III etc.
        r"\b(B\.?\s*Tech\.?\s*(?:I|II|III|IV))\b",

        # Academic year range
        r"\b(20\d{2}\s*[-/]\s*20\d{2})\b",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            value = _clean_ocr_value(
                match.group(match.lastindex)
            )

            if value:

                replacements = {
                    "first year": "First Year",
                    "second year": "Second Year",
                    "third year": "Third Year",
                    "fourth year": "Fourth Year",
                    "1st year": "1st Year",
                    "2nd year": "2nd Year",
                    "3rd year": "3rd Year",
                    "4th year": "4th Year",
                    "fy": "First Year",
                    "sy": "Second Year",
                    "ty": "Third Year",
                    "final year": "Final Year",
                }

                return replacements.get(
                    value.lower(),
                    value
                )

    # --------------------------------------------------------
    # FALLBACK:
    # Explicit calendar year such as 2026
    # --------------------------------------------------------

    match = re.search(
        r"\b(20\d{2})\b",
        text
    )

    if match:
        return match.group(1)

    # --------------------------------------------------------
    # FALLBACK:
    # Event names such as RIT-HACKATHON 2K26
    # --------------------------------------------------------

    match = re.search(
        r"\b2K(\d{2})\b",
        text,
        re.IGNORECASE
    )

    if match:
        return "20" + match.group(1)

    return ""


def extract_date_value(text):
    """Extract date from different certificate formats."""

    if not text:
        return ""

    # --------------------------------------------------------
    # Normalize common OCR mistakes in ordinal dates.
    #
    # Example:
    # 24” and 25° March, 2026
    # becomes approximately:
    # 24th and 25th March, 2026
    # --------------------------------------------------------

    normalized_text = text

    normalized_text = re.sub(
        r"(\d{1,2})\s*[\"”°º]",
        r"\1th",
        normalized_text
    )

    # OCR may also produce strange spaces
    normalized_text = re.sub(
        r"\s+",
        " ",
        normalized_text
    )

    month_pattern = (
        r"(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December)"
    )

    patterns = [

        # Date: 20/09/2026
        r"(?:Date|Dated)\s*[:\-]?\s*"
        r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",

        # 20/09/2026
        r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\b",

        # Date: 24th and 25th March, 2026
        r"(?:Date|Dated)\s*[:\-]?\s*"
        r"(\d{1,2}(?:st|nd|rd|th)?\s+"
        r"(?:and|&)\s+"
        r"\d{1,2}(?:st|nd|rd|th)?\s+"
        + month_pattern +
        r"(?:,)?\s+\d{4})",

        # 24th and 25th March, 2026
        r"\b(\d{1,2}(?:st|nd|rd|th)?\s+"
        r"(?:and|&)\s+"
        r"\d{1,2}(?:st|nd|rd|th)?\s+"
        + month_pattern +
        r"(?:,)?\s+\d{4})\b",

        # 24th-25th March 2026
        r"\b(\d{1,2}(?:st|nd|rd|th)?\s*"
        r"(?:-|–|—|to)\s*"
        r"\d{1,2}(?:st|nd|rd|th)?\s+"
        + month_pattern +
        r"(?:,)?\s+\d{4})\b",

        # 24th March, 2026
        r"\b(\d{1,2}(?:st|nd|rd|th)?\s+"
        + month_pattern +
        r"(?:,)?\s+\d{4})\b",

        # March 24, 2026
        r"\b(("
        + month_pattern +
        r")\s+\d{1,2}(?:st|nd|rd|th)?"
        r"(?:,\s*|\s+)\d{4})\b",

        # 24th and 25th of March, 2026
        r"\b(\d{1,2}(?:st|nd|rd|th)?\s+"
        r"(?:and|&)\s+"
        r"\d{1,2}(?:st|nd|rd|th)?\s+of\s+"
        + month_pattern +
        r"(?:,)?\s+\d{4})\b",

        # 24th and 25th March
        r"\b(\d{1,2}(?:st|nd|rd|th)?\s+"
        r"(?:and|&)\s+"
        r"\d{1,2}(?:st|nd|rd|th)?\s+"
        + month_pattern +
        r")\b",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            normalized_text,
            re.IGNORECASE
        )

        if match:

            date_value = _clean_ocr_value(
                match.group(1)
            )

            if date_value:
                return date_value

    return ""

def extract_time_value(text):
    """Extract event time."""
    if not text:
        return ""

    patterns = [
        r"\bTime\s*[:\-]?\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)",
        r"\bat\s+(\d{1,2}[:.]\d{2}\s*(?:AM|PM))\b",
        r"\b(\d{1,2}[:.]\d{2}\s*(?:AM|PM))\b",
        r"\b(\d{1,2}:\d{2})\b",
        r"\b(\d{1,2}\s*(?:AM|PM))\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            value = _clean_ocr_value(match.group(1))

            if value:
                return value

    return ""


def extract_event_type(text):
    """Extract event type."""
    if not text:
        return ""

    event_types = [
        ("Hackathon", r"\bhackathon\b"),
        ("Workshop", r"\bworkshop\b"),
        ("Webinar", r"\bwebinar\b"),
        ("Seminar", r"\bseminar\b"),
        ("Conference", r"\bconference\b"),
        ("Bootcamp", r"\bboot\s*camp\b|\bbootcamp\b"),
        ("Training", r"\btraining\b"),
        ("Internship", r"\binternship\b"),
        ("Quiz Competition", r"\bquiz\s+competition\b"),
        ("Coding Competition", r"\bcoding\s+competition\b|\bcoding\s+contest\b"),
        ("Technical Competition", r"\btechnical\s+competition\b|\btechnical\s+contest\b"),
        ("Project Exhibition", r"\bproject\s+exhibition\b|\bproject\s+expo\b"),
        ("Paper Presentation", r"\bpaper\s+presentation\b"),
        ("Poster Presentation", r"\bposter\s+presentation\b"),
        ("Competition", r"\bcompetition\b|\bcontest\b"),
        ("Presentation", r"\bpresentation\b"),
    ]

    for event_type, pattern in event_types:
        if re.search(pattern, text, re.IGNORECASE):
            return event_type

    if re.search(
        r"\bcertificate\s+of\s+merit\b",
        text,
        re.IGNORECASE
    ) and re.search(
        r"\b(?:secured|won|stood)\s+\d+(?:st|nd|rd|th)?\s+"
        r"(?:place|position)\b",
        text,
        re.IGNORECASE
    ):
        return "Competition"

    return ""


def extract_college_name(text):
    """Extract college/institute/university name."""
    if not text:
        return ""

    patterns = [
        r"(?:College|Institute|University|Institution)\s*"
        r"(?:Name)?\s*[:\-]\s*"
        r"([A-Za-z0-9 .,&'()\-]{5,150})"
        r"(?=\s+(?:Event|Activity|Date|Time|Year|Branch|Department|"
        r"PRN|Roll|Participant|Student|Name|Location|Venue)\b|$)",

        r"\b(Rajarambapu\s+Institute\s+of\s+Technology)\b",

        r"\b(Sathyabama\s+Institute\s+of\s+Science\s+and\s+Technology)\b",

        r"\b([A-Z][A-Za-z0-9&'.-]*(?:\s+[A-Z][A-Za-z0-9&'.()-]*){1,10}"
        r"\s+(?:Institute|University|College)"
        r"(?:\s+of\s+[A-Za-z0-9&'.() -]+)?)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            college = _clean_ocr_value(match.group(1))

            college = re.split(
                r"\s*,\s*(?:Rajaramnagar|Pune|Mumbai|Kolhapur|"
                r"Sangli|Chennai|Bengaluru|Bangalore)\b",
                college,
                maxsplit=1,
                flags=re.IGNORECASE
            )[0]

            if college:
                return college

    return ""

def extract_all_certificate_fields(text):
    """Extract all required certificate fields."""

    return {
        "year": extract_year(text),
        "date": extract_date_value(text),
        "time": extract_time_value(text),
        "prn": extract_prn(text),
        "participant_name": extract_participant_name(text),
        "branch": extract_branch(text),
        "event_type": extract_event_type(text),
        "college_name": extract_college_name(text),
    }

@login_required
def upload_certificate(request):

    student = Student.objects.filter(
        email=request.user.email
    ).first()

    if student is None:
        return render(
            request,
            'accounts/upload_certificate.html',
            {
                'error': 'No student record linked to this account.'
            }
        )

    student_activities = Activity.objects.filter(
        student=student
    ).order_by('-start_date')

    if request.method == 'POST':

        activity_id = request.POST.get('activity_id')
        uploaded_file = request.FILES.get('certificate_file')

        activity = Activity.objects.filter(
            activity_id=activity_id,
            student=student
        ).first()

        if activity is None or uploaded_file is None:
            return render(
                request,
                'accounts/upload_certificate.html',
                {
                    'student_activities': student_activities,
                    'error': 'Please select an activity and choose a file.'
                }
            )

        # ----------------------------------------------------
        # SAVE FILE TEMPORARILY
        # ----------------------------------------------------

        temp_dir = os.path.join(
            settings.MEDIA_ROOT,
            'temp_certificates'
        )

        os.makedirs(
            temp_dir,
            exist_ok=True
        )

        temp_path = os.path.join(
            temp_dir,
            uploaded_file.name
        )

        with open(
            temp_path,
            'wb+'
        ) as destination:

            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        # ----------------------------------------------------
        # CHECK FILE TYPE
        # ----------------------------------------------------

        file_extension = os.path.splitext(
            uploaded_file.name
        )[1].lower()

        allowed_extensions = [
            ".pdf",
            ".jpg",
            ".jpeg",
            ".png",
        ]

        if file_extension not in allowed_extensions:

            os.remove(temp_path)

            return render(
                request,
                'accounts/upload_certificate.html',
                {
                    'student_activities': student_activities,
                    'error':
                        'Only PDF, JPG, JPEG and PNG files are allowed.'
                }
            )

        # ----------------------------------------------------
        # EXTRACT CERTIFICATE TEXT
        # ----------------------------------------------------

        try:

            extracted_text = extract_certificate_text(
                temp_path
            )

        except Exception as error:

            if os.path.exists(temp_path):
                os.remove(temp_path)

            return render(
                request,
                'accounts/upload_certificate.html',
                {
                    'student_activities': student_activities,
                    'error':
                        f'Unable to read certificate: {str(error)}'
                }
            )

        # ----------------------------------------------------
        # EXTRACT ALL 8 CERTIFICATE FIELDS
        # ----------------------------------------------------

        extracted_fields = extract_all_certificate_fields(
            extracted_text
        )

        extracted_prn = extracted_fields.get(
            "prn",
            ""
        )

        extracted_name = extracted_fields.get(
            "participant_name",
            ""
        )

        extracted_branch = extracted_fields.get(
            "branch",
            ""
        )

        extracted_year = extracted_fields.get(
            "year",
            ""
        )

        extracted_date = extracted_fields.get(
            "date",
            ""
        )

        extracted_time = extracted_fields.get(
            "time",
            ""
        )

        extracted_event_type = extracted_fields.get(
            "event_type",
            ""
        )

        extracted_college_name = extracted_fields.get(
            "college_name",
            ""
        )

        # ----------------------------------------------------
        # PRN CHECK
        # ----------------------------------------------------

        if (
            extracted_prn
            and str(extracted_prn).strip()
            != str(student.prn).strip()
        ):

            if os.path.exists(temp_path):
                os.remove(temp_path)

            return render(
                request,
                'accounts/upload_certificate.html',
                {
                    'student_activities': student_activities,

                    'error':
                        f'Upload cancelled: The PRN on the certificate '
                        f'({extracted_prn}) does not match your account '
                        f'({student.prn}).'
                }
            )

        # ----------------------------------------------------
        # SAVE PENDING CERTIFICATE DATA IN SESSION
        # ----------------------------------------------------

        request.session['pending_certificate'] = {

            'temp_filename': uploaded_file.name,

            'activity_id': activity_id,

            'year': extracted_year,

            'date': extracted_date,

            'time': extracted_time,

            'prn': extracted_prn,

            'participant_name': extracted_name,

            'branch': extracted_branch,

            'event_type': extracted_event_type,

            'college_name': extracted_college_name,
        }

        # ----------------------------------------------------
        # SHOW PREVIEW
        # ----------------------------------------------------

        return render(
            request,
            'accounts/upload_certificate.html',
            {
                'student_activities': student_activities,

                'extracted_name':
                    extracted_name or 'Not found',

                'extracted_prn':
                    extracted_prn or 'Not found',

                'extracted_branch':
                    extracted_branch or 'Not found',

                'extracted_year':
                    extracted_year or 'Not found',

                'extracted_date':
                    extracted_date or 'Not found',

                'extracted_time':
                    extracted_time or 'Not found',

                'extracted_event_type':
                    extracted_event_type or 'Not found',

                'extracted_college_name':
                    extracted_college_name or 'Not found',

                'show_confirm': True,
            }
        )

    # --------------------------------------------------------
    # NORMAL PAGE LOAD
    # --------------------------------------------------------

    return render(
        request,
        'accounts/upload_certificate.html',
        {
            'student_activities': student_activities,
        }
    )


@login_required
def confirm_certificate(request):

    student = Student.objects.filter(
        email=request.user.email
    ).first()

    if student is None or request.method != 'POST':
        return redirect('upload_certificate')

    pending = request.session.get(
        'pending_certificate'
    )

    if not pending:
        return redirect('upload_certificate')

    temp_path = os.path.join(
        settings.MEDIA_ROOT,
        'temp_certificates',
        pending['temp_filename']
    )

    if not os.path.exists(temp_path):

        del request.session[
            'pending_certificate'
        ]

        return render(
            request,
            'accounts/upload_certificate.html',
            {
                'student_activities':
                    Activity.objects.filter(
                        student=student
                    ).order_by('-start_date'),

                'error':
                    'Certificate file expired, please upload again.'
            }
        )

    activity = Activity.objects.filter(
        activity_id=pending['activity_id'],
        student=student
    ).first()

    if activity is None:

        os.remove(temp_path)

        del request.session[
            'pending_certificate'
        ]

        return redirect(
            'upload_certificate'
        )

    # --------------------------------------------------------
    # MOVE CERTIFICATE TO FINAL STORAGE
    # --------------------------------------------------------

    save_dir = os.path.join(
        settings.MEDIA_ROOT,
        'certificates'
    )

    os.makedirs(
        save_dir,
        exist_ok=True
    )

    final_path = os.path.join(
        save_dir,
        pending['temp_filename']
    )

    os.replace(
        temp_path,
        final_path
    )

    # --------------------------------------------------------
    # CREATE CERTIFICATE DATABASE RECORD
    # --------------------------------------------------------

    certificate_file_name = (
        f'certificates/{pending["temp_filename"]}'
    )

    certificate = Certificate.objects.create(
        student=student,
        activity=activity,
        certificate_file=certificate_file_name,
        verification_status='Pending',
        upload_date=timezone.now(),
    )

    # --------------------------------------------------------
    # PREPARE EXTRACTED OCR DATA
    # --------------------------------------------------------

    excel_data = {
        'year': pending.get('year', ''),
        'date': pending.get('date', ''),
        'time': pending.get('time', ''),
        'prn': pending.get('prn', ''),
        'participant_name': pending.get(
            'participant_name',
            ''
        ),
        'branch': pending.get('branch', ''),
        'event_type': pending.get(
            'event_type',
            ''
        ),
        'college_name': pending.get(
            'college_name',
            ''
        ),
    }

    # --------------------------------------------------------
    # SAVE OCR MAPPING FOR FACULTY
    # --------------------------------------------------------

    ocr_directory = os.path.join(
        settings.MEDIA_ROOT,
        'excel'
    )

    os.makedirs(
        ocr_directory,
        exist_ok=True
    )

    ocr_mapping_path = os.path.join(
        ocr_directory,
        'certificate_extracted_data.json'
    )

    certificate_extracted_data = {}

    if os.path.exists(ocr_mapping_path):

        try:

            with open(
                ocr_mapping_path,
                'r',
                encoding='utf-8'
            ) as json_file:

                certificate_extracted_data = json.load(
                    json_file
                )

        except (
            json.JSONDecodeError,
            OSError
        ):

            certificate_extracted_data = {}

    certificate_extracted_data[
        certificate_file_name
    ] = excel_data

    try:

        with open(
            ocr_mapping_path,
            'w',
            encoding='utf-8'
        ) as json_file:

            json.dump(
                certificate_extracted_data,
                json_file,
                ensure_ascii=False,
                indent=4
            )

    except OSError as error:

        print(
            "Certificate OCR mapping save error:",
            error
        )

    # --------------------------------------------------------
    # ADD EXTRACTED DATA TO EXCEL
    # --------------------------------------------------------

    append_certificate_to_excel(
        excel_data
    )

    # --------------------------------------------------------
    # CLEAR PENDING SESSION
    # --------------------------------------------------------

    del request.session[
        'pending_certificate'
    ]

    return redirect(
        'my_certificates'
    )

@login_required
def cancel_certificate(request):

    pending = request.session.pop(
        'pending_certificate',
        None
    )

    if pending:

        temp_path = os.path.join(
            settings.MEDIA_ROOT,
            'temp_certificates',
            pending['temp_filename']
        )

        if os.path.exists(temp_path):
            os.remove(temp_path)

    return redirect('upload_certificate')

@login_required
def download_certificate_excel(request):
    """Download the currently extracted certificate details as Excel."""

    pending_certificate = request.session.get('pending_certificate')

    if not pending_certificate:
        messages.error(
            request,
            'Extracted certificate data is not available.'
        )
        return redirect('upload_certificate')

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Certificate Details"

    # 8 extracted certificate fields
    headers = [
        "Year",
        "Date",
        "Time",
        "PRN",
        "Participant Name",
        "Branch",
        "Event Type",
        "College Name",
    ]

    worksheet.append(headers)

    # Put the ACTUAL extracted values into Excel
    worksheet.append([
        pending_certificate.get("year", ""),
        pending_certificate.get("date", ""),
        pending_certificate.get("time", ""),
        pending_certificate.get("prn", ""),
        pending_certificate.get("participant_name", ""),
        pending_certificate.get("branch", ""),
        pending_certificate.get("event_type", ""),
        pending_certificate.get("college_name", ""),
    ])

    # Column widths
    widths = [15, 30, 18, 18, 30, 25, 20, 40]

    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=index).column_letter
        ].width = width

    response = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    response["Content-Disposition"] = (
        'attachment; filename="Certificate_Extracted_Data.xlsx"'
    )

    workbook.save(response)

    return response

@login_required
def extract_certificate(request):

    if request.method != "POST":
        return JsonResponse({
            "success": False,
            "error": "Invalid request method."
        })

    student = Student.objects.filter(
        email=request.user.email
    ).first()

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

    file_extension = os.path.splitext(
        uploaded_file.name
    )[1].lower()

    allowed_extensions = [
        ".pdf",
        ".jpg",
        ".jpeg",
        ".png"
    ]

    if file_extension not in allowed_extensions:
        return JsonResponse({
            "success": False,
            "error": "Only PDF, JPG, JPEG and PNG files are allowed."
        })

    temp_dir = os.path.join(
        settings.MEDIA_ROOT,
        "temp_certificate_extraction"
    )

    os.makedirs(
        temp_dir,
        exist_ok=True
    )

    temp_filename = (
        f"temp_{request.user.id}_"
        f"{int(time.time() * 1000)}"
        f"{file_extension}"
    )

    temp_path = os.path.join(
        temp_dir,
        temp_filename
    )

    try:

        # ---------------------------------------------------------
        # SAVE UPLOADED FILE TEMPORARILY
        # ---------------------------------------------------------

        with open(temp_path, "wb+") as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        # ---------------------------------------------------------
        # EXTRACT TEXT USING IMPROVED OCR
        # ---------------------------------------------------------

        extracted_text = extract_certificate_text(
            temp_path
        )

        extracted_text = clean_certificate_text(
            extracted_text
        )

        if not extracted_text:
            return JsonResponse({
                "success": False,
                "error": (
                    "No readable text was found in "
                    "the certificate."
                )
            })

        # ---------------------------------------------------------
        # EXTRACT ALL 8 CERTIFICATE FIELDS
        # ---------------------------------------------------------

        try:

            extracted_fields = extract_all_certificate_fields(
                extracted_text
            )

        except Exception as error:

            print(
                "Certificate field extraction error:",
                error
            )

            extracted_fields = {}

        # ---------------------------------------------------------
        # GET INDIVIDUAL VALUES
        # ---------------------------------------------------------

        extracted_year = str(
            extracted_fields.get("year", "")
        ).strip()

        extracted_date = str(
            extracted_fields.get("date", "")
        ).strip()

        extracted_time = str(
            extracted_fields.get("time", "")
        ).strip()

        extracted_prn = str(
            extracted_fields.get("prn", "")
        ).strip()

        extracted_name = str(
            extracted_fields.get(
                "participant_name",
                ""
            )
        ).strip()

        extracted_branch = str(
            extracted_fields.get("branch", "")
        ).strip()

        extracted_event_type = str(
            extracted_fields.get(
                "event_type",
                ""
            )
        ).strip()

        extracted_college_name = str(
            extracted_fields.get(
                "college_name",
                ""
            )
        ).strip()

        # ---------------------------------------------------------
        # PRN VALIDATION
        # ---------------------------------------------------------

        if extracted_prn:

            student_prn = str(
                student.prn
            ).strip()

            if extracted_prn != student_prn:

                return JsonResponse({
                    "success": False,
                    "error": (
                        "The PRN found on the certificate "
                        "does not match your account."
                    )
                })

        # ---------------------------------------------------------
        # RETURN ALL 8 FIELDS
        # ---------------------------------------------------------

        return JsonResponse({

            "success": True,

            "year": extracted_year,

            "date": extracted_date,

            "time": extracted_time,

            "prn": extracted_prn,

            "participant_name": extracted_name,

            "branch": extracted_branch,

            "event_type": extracted_event_type,

            "college_name": extracted_college_name,

        })

    except Exception as error:

        print(
            "Certificate extraction error:",
            error
        )

        return JsonResponse({
            "success": False,
            "error": (
                f"Unable to process certificate: "
                f"{str(error)}"
            )
        })

    finally:

        # ---------------------------------------------------------
        # DELETE TEMPORARY FILE
        # ---------------------------------------------------------

        try:

            if os.path.exists(temp_path):
                os.remove(temp_path)

        except Exception as error:

            print(
                "Temporary certificate cleanup error:",
                error
            )

def download_certificate_excel(request):
    """Download the logged-in student's certificates as an Excel file."""

    if not request.user.is_authenticated:
        return redirect("login")

    try:
        student = Student.objects.get(email=request.user.email)
    except Student.DoesNotExist:
        return HttpResponse("Student record not found.", status=404)

    certificates = Certificate.objects.filter(
        student=student
    ).select_related("activity").order_by("-upload_date")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "My Certificates"

    # Excel headings
    headers = [
        "Sr. No.",
        "Activity Name",
        "Uploaded On",
        "Verification Status",
    ]

    worksheet.append(headers)

    # Certificate data
    for index, certificate in enumerate(certificates, start=1):
        worksheet.append([
            index,
            certificate.activity.activity_name
            if certificate.activity else "",
            certificate.upload_date.strftime("%d-%m-%Y")
            if certificate.upload_date else "",
            certificate.verification_status or "",
        ])

    # Make columns wider
    worksheet.column_dimensions["A"].width = 10
    worksheet.column_dimensions["B"].width = 35
    worksheet.column_dimensions["C"].width = 18
    worksheet.column_dimensions["D"].width = 22

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    response["Content-Disposition"] = (
        'attachment; filename="My_Certificates.xlsx"'
    )

    workbook.save(response)

    return response

def download_extracted_certificate_excel(request):
    """Download the 8 extracted certificate fields as Excel."""

    if not request.user.is_authenticated:
        return redirect("login")

    pending_certificate = request.session.get("pending_certificate")

    if not pending_certificate:
        return HttpResponse(
            "Extracted certificate data is not available.",
            status=404
        )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Extracted Certificate"

    # 8 required OCR fields
    headers = [
        "Year",
        "Date",
        "Time",
        "PRN",
        "Participant Name",
        "Branch",
        "Event Type",
        "College Name",
    ]

    worksheet.append(headers)

    # Add the actual extracted values
    worksheet.append([
        pending_certificate.get("year", ""),
        pending_certificate.get("date", ""),
        pending_certificate.get("time", ""),
        pending_certificate.get("prn", ""),
        pending_certificate.get("participant_name", ""),
        pending_certificate.get("branch", ""),
        pending_certificate.get("event_type", ""),
        pending_certificate.get("college_name", ""),
    ])

    # Column widths
    column_widths = {
        "A": 15,
        "B": 30,
        "C": 18,
        "D": 18,
        "E": 30,
        "F": 25,
        "G": 20,
        "H": 40,
    }

    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    response = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    response["Content-Disposition"] = (
        'attachment; filename="Certificate_Extracted_Data.xlsx"'
    )

    workbook.save(response)

    return response

@login_required
def download_certificate_excel_for_faculty(
    request,
    certificate_id
):
    """
    Download the exact 8 certificate fields extracted
    during the student's upload.

    The function first checks the saved OCR mapping.
    For older certificates without a saved mapping,
    it falls back to OCR on the certificate file.
    """

    # ---------------------------------------------------------
    # Get the selected certificate
    # ---------------------------------------------------------

    certificate = Certificate.objects.filter(
        certificate_id=certificate_id
    ).select_related(
        "student",
        "activity"
    ).first()

    if certificate is None:
        return HttpResponse(
            "Certificate not found.",
            status=404
        )

    # ---------------------------------------------------------
    # Read OCR values saved during certificate confirmation
    # ---------------------------------------------------------

    ocr_directory = os.path.join(
        settings.MEDIA_ROOT,
        "excel"
    )

    ocr_mapping_path = os.path.join(
        ocr_directory,
        "certificate_extracted_data.json"
    )

    extracted_data = {}

    if os.path.exists(ocr_mapping_path):
        try:
            with open(
                ocr_mapping_path,
                "r",
                encoding="utf-8"
            ) as json_file:
                certificate_extracted_data = json.load(
                    json_file
                )

            certificate_key = str(
                certificate.certificate_file
            )

            extracted_data = (
                certificate_extracted_data.get(
                    certificate_key,
                    {}
                )
            )

        except (
            json.JSONDecodeError,
            OSError
        ):
            extracted_data = {}

    # ---------------------------------------------------------
    # Fallback for older certificates
    # ---------------------------------------------------------

    if not extracted_data:

        certificate_path = os.path.join(
            settings.MEDIA_ROOT,
            str(certificate.certificate_file)
        )

        if os.path.exists(certificate_path):
            try:
                certificate_text = (
                    extract_certificate_text(
                        certificate_path
                    )
                )

                extracted_data = (
                    extract_all_certificate_fields(
                        certificate_text
                    )
                )

            except Exception:
                extracted_data = {}

    # ---------------------------------------------------------
    # Create Excel workbook
    # ---------------------------------------------------------

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Certificate Details"

    # ---------------------------------------------------------
    # Required 8 fields
    # ---------------------------------------------------------

    headers = [
        "Year",
        "Date",
        "Time",
        "PRN",
        "Participant Name",
        "Branch",
        "Event Type",
        "College Name",
    ]

    worksheet.append(headers)

    # ---------------------------------------------------------
    # Add extracted values
    # ---------------------------------------------------------

    worksheet.append([
        extracted_data.get("year", ""),
        extracted_data.get("date", ""),
        extracted_data.get("time", ""),
        extracted_data.get("prn", ""),
        extracted_data.get("participant_name", ""),
        extracted_data.get("branch", ""),
        extracted_data.get("event_type", ""),
        extracted_data.get("college_name", ""),
    ])

    # ---------------------------------------------------------
    # Column widths
    # ---------------------------------------------------------

    column_widths = {
        "A": 15,
        "B": 30,
        "C": 18,
        "D": 18,
        "E": 30,
        "F": 25,
        "G": 25,
        "H": 40,
    }

    for column, width in column_widths.items():
        worksheet.column_dimensions[
            column
        ].width = width

    # ---------------------------------------------------------
    # Return Excel file
    # ---------------------------------------------------------

    response = HttpResponse(
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    response["Content-Disposition"] = (
        f'attachment; '
        f'filename="Certificate_'
        f'{certificate_id}_Extracted_Data.xlsx"'
    )

    workbook.save(response)

    return response

@login_required
def delete_student(request, student_id):
    if request.user.role != 'ADMIN':
        return redirect('dashboard_redirect')

    if request.method == 'POST':
        with transaction.atomic():
            student = get_object_or_404(Student, student_id=student_id)

            # Delete student's Django login account
            User.objects.filter(
                email=student.email,
                role='STUDENT'
            ).delete()

            # Delete student record
            # Related activities, certificates and leave applications
            # are deleted because of on_delete=models.CASCADE
            student_name = student.full_name
            student.delete()

        messages.success(
            request,
            f'Student "{student_name}" deleted successfully.'
        )

    return redirect('manage_students')


@login_required
def delete_faculty(request, faculty_id):
    if request.user.role != 'ADMIN':
        return redirect('dashboard_redirect')

    if request.method == 'POST':
        with transaction.atomic():
            faculty = get_object_or_404(
                Faculty,
                faculty_id=faculty_id
            )

            # Delete faculty's Django login account
            User.objects.filter(
                employee_id=faculty.employee_id,
                role='FACULTY'
            ).delete()

            # Delete faculty record
            faculty_name = faculty.full_name
            faculty.delete()

        messages.success(
            request,
            f'Faculty "{faculty_name}" deleted successfully.'
        )

    return redirect('manage_faculty')

