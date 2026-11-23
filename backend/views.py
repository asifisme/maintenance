import code
import os
import base64
import logging
import qrcode 
import json
import uuid
import calendar
from datetime import datetime, date 
from urllib import request
from io import BytesIO 
from django.core.files import File 
from django.urls import reverse
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.models import User
from django.contrib.auth.forms import AdminPasswordChangeForm
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.views.generic import ListView
from django.views.decorators.http import require_GET, require_POST
from django.core.cache import cache
from django.core.paginator import Paginator, PageNotAnInteger, EmptyPage
from django.http import JsonResponse, HttpResponse
from django.conf import settings
from django.utils.text import slugify
from django.views.generic import ListView, CreateView, UpdateView, DetailView
from django.db.models import Q
from datetime import date, datetime, timedelta
from decimal import Decimal
from django.utils import timezone
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from decimal import Decimal


from django.http import HttpResponse
from datetime import datetime
from PIL import Image
from io import BytesIO
from django.db.models import Min
from django.views.generic import ListView, CreateView, UpdateView, DetailView 
from .common_func import checkUserPermission
from backend.forms import UserCreateForm 


from backend.models import (
    LoginLog, BackendMenu, UserMenuPermission, WebImages, SiteSettings,
    Division, SubDivision, Section, Building, EquipmentType, EquipmentTypeData, Equipment,
    Ticket, TicketTaskItem, TicketComment, TicketActivityLog, AssignTechnician, Skill, Technician, AssignActivities,
    ScheduledMaintenance, CriticalAlert, Maintenance, MaintenanceRecord, MaintenanceAttachment,
    EquipmentComponents, MaintenanceComponent, EquipmentComponentPreset, 
)
from django.db.models import Count

from backend.forms import (
    CustomUserLoginForm,
    MaintenanceForm,
)


logger = logging.getLogger(__name__)


# Format timezone-aware datetimes to local time for display/export
def format_local_time(dt_obj, fmt='%H:%M'):
    """Return local time formatted with fmt or '--:--' if missing."""
    if not dt_obj:
        return '--:--'
    local_dt = timezone.localtime(dt_obj) if timezone.is_aware(dt_obj) else dt_obj
    return local_dt.strftime(fmt)

def format_local_datetime(dt_obj):
    """Return local datetime string or empty string if missing."""
    if not dt_obj:
        return ''
    local_dt = timezone.localtime(dt_obj) if timezone.is_aware(dt_obj) else dt_obj
    return local_dt.strftime('%d-%b-%Y %H:%M:%S')

def paginate_data(request, page_num, data_list):
    if hasattr(data_list, 'ordered') and not data_list.ordered:
        data_list = data_list.order_by('id')
    try:
        items_per_page = int(request.GET.get('per_page', 10))
        if items_per_page <= 0:
            items_per_page = 10
    except (ValueError, TypeError):
        items_per_page = 10
    max_pages = 5
    paginator = Paginator(data_list, items_per_page)
    last_page_number = paginator.num_pages

    try:
        data_list = paginator.page(page_num)
    except PageNotAnInteger:
        data_list = paginator.page(1)
    except EmptyPage:
        data_list = paginator.page(paginator.num_pages)

    current_page = data_list.number
    start_page = max(current_page - int(max_pages / 2), 1)
    end_page = start_page + max_pages

    if end_page > last_page_number:
        end_page = last_page_number + 1
        start_page = max(end_page - max_pages, 1)

    paginator_list = range(start_page, end_page)

    return data_list, paginator_list, last_page_number


def _resolve_menu_url(menu_url: str):
    if not menu_url:
        return ""
    if "/" in menu_url:
        return menu_url
    try:
        return reverse("backend:menu_wise_dashboard", args=[menu_url])
    except Exception:
        return menu_url


def _score_menu(menu, ql: str) -> int:
    name = (menu.menu_name or "").lower()
    module = (menu.module_name or "").lower()
    desc = (menu.menu_description or "").lower()

    score = 0

    if name == ql:
        score += 1000
    if name.startswith(ql):
        score += 300
    if ql in name:
        score += 200
        pos = name.find(ql)
        score += max(0, 100 - min(pos, 100))

    if module == ql:
        score += 120
    elif module.startswith(ql):
        score += 80
    elif ql in module:
        score += 40

    if ql in desc:
        score += 10

    score += max(0, 30 - abs(len(name) - len(ql)))

    return score


def serve_optimized_image(request, unique_key):
    try:
        width = request.GET.get("width")
        quality = request.GET.get("quality")
        path = request.GET.get("image_path")

        try:
            width = int(width) if width else None
        except ValueError:
            width = None

        try:
            quality = int(quality) if quality else 80
        except ValueError:
            quality = 80

        if quality > 100:
            quality = 100

        cache_key = f"photo_{unique_key}_{width or 'auto'}_{quality}_{path or 'default'}"
        cached_image = cache.get(cache_key)
        if cached_image:
            image_bytes = base64.b64decode(cached_image)
            response = HttpResponse(image_bytes, content_type="image/webp")
            response["Cache-Control"] = "public, max-age=2592000, immutable"
            return response

        if unique_key:
            if unique_key != "no_image":
                img_obj = get_object_or_404(WebImages, unique_key=unique_key)
                image_path = img_obj.image.path
            elif unique_key == "no_image" and request.GET.get("image_path"):
                image_path = request.GET.get("image_path")
                if image_path == "logo" or image_path == "favicon":
                    site_settings = SiteSettings.objects.first()
                    if image_path == "logo":
                        image_path = site_settings.logo.path if site_settings.logo else os.path.join(settings.STATICFILES_DIRS[0], "images/default_logo.png")
                    elif image_path == "favicon":
                        image_path = site_settings.favicon.path if site_settings.favicon else os.path.join(settings.STATICFILES_DIRS[0], "images/default_favicon.png")
                elif image_path.startswith(settings.MEDIA_URL):
                    image_path = image_path.replace(settings.MEDIA_URL, settings.MEDIA_ROOT + "/")
            else:
                image_path = os.path.join(settings.STATICFILES_DIRS[0], "images/no_image.png")
        else:
            image_path = os.path.join(settings.STATICFILES_DIRS[0], "images/no_image.png")

        img = Image.open(image_path)

        if width:
            ratio = width / float(img.width)
            height = int(img.height * ratio)
            img = img.resize((width, height), Image.LANCZOS)

        buffer = BytesIO()
        img.save(buffer, format="WEBP", quality=quality, optimize=True)
        image_bytes = buffer.getvalue()

        cache.set(cache_key, base64.b64encode(image_bytes).decode("ascii"), timeout=86400)

        base_name = slugify(unique_key)
        filename = f"{base_name}.webp"
        response = HttpResponse(image_bytes, content_type="image/webp")
        response["Content-Disposition"] = f'inline; filename="{filename}"'
        response["Cache-Control"] = "public, max-age=2592000, immutable"
        return response

    except Exception:
        # Redirect to "no_image" instead of returning a string
        no_image_url = reverse('backend:serve_optimized_image', kwargs={'unique_key': 'no_image'})
        no_image_url = f"{no_image_url}?width=300&quality=80"
        return redirect(no_image_url)


@require_GET
def search_backend_menus(request):
    if not request.user.is_authenticated:
        return JsonResponse({"status": False, "auth_required": True, "data": []}, status=401)

    q = (request.GET.get("q") or "").strip()
    if not q:
        return JsonResponse({"status": True, "data": []})

    try:
        limit = int(request.GET.get("limit", 10))
    except ValueError:
        limit = 10
    limit = max(1, min(limit, 50))

    perms = (
        UserMenuPermission.objects.filter(
            user_id=request.user.id, can_view=True, is_active=True, deleted=False, menu__is_active=True,
        )
        .select_related("menu")
        .order_by("menu__id")
    )

    like = Q(menu__menu_name__icontains=q) | Q(menu__parent__menu_name__icontains=q)

    ql = q.lower()
    seen = set()
    ranked = []

    for perm in perms.filter(like):
        m = perm.menu
        if m.id in seen:
            continue
        seen.add(m.id)
        ranked.append((_score_menu(m, ql), m))

    ranked.sort(key=lambda t: (-t[0], len((t[1].menu_name or "")), (t[1].menu_name or "").lower()))

    results = []
    for _, m in ranked[: limit]:
        url = _resolve_menu_url(m.menu_url)
        icon = m.menu_icon or "fa-solid fa-circle"
        title = m.menu_name or ""
        # description = m.menu_description or ""

        parent_menus = []
        current_menu = m
        while current_menu.parent:
            parent_menus.append(current_menu.parent.menu_name)
            current_menu = current_menu.parent
        parent_menus.reverse()

        description = " > ".join(parent_menus) if parent_menus else ""

        results.append(
            {
                "name": title,
                "description": description,
                "icon": icon,
                "url": url,
                "module": m.module_name or "",
            }
        )

    return JsonResponse({"status": True, "count": len(results), "data": results})



def backend_login(request):
    if request.user.is_authenticated:
        return redirect('backend:backend_logout')

    form = CustomUserLoginForm(request.POST or None)

    if request.method == 'POST' and form.is_valid():
        username = form.cleaned_data.get('username')
        password = form.cleaned_data.get('password')

        user_ip = request.META.get('HTTP_X_FORWARDED_FOR') or request.META.get('REMOTE_ADDR')

        user = User.objects.filter(username=username).first()

        if user:
            authenticated_user = authenticate(request, username=username, password=password)

            if authenticated_user is not None:
                login(request, authenticated_user)
                LoginLog.objects.create(user_id=user.id, username=username, login_ip=user_ip, login_status=True)

                if user.is_superuser:
                    menu_list = BackendMenu.objects.filter(is_active=True)
                    for menu in menu_list:
                        UserMenuPermission.objects.update_or_create(
                            user_id=user.id,
                            menu_id=menu.id,
                            defaults={
                                'can_view': True,
                                'can_add': True,
                                'can_update': True,
                                'can_delete': True,
                                'is_active': True,
                                'created_by_id': request.user.id,
                            }
                        )

                next_url = '/'
                return redirect(next_url)

        LoginLog.objects.create(username=username, login_ip=user_ip, login_status=False)
        messages.error(request, "Invalid username or password.")
    context = {
        'form': form
    }
    return render(request, 'backend_login.html', context)


@login_required
def backend_logout(request):
    LoginLog.objects.create(
        user_id=request.user.id,
        username=request.user.username,
        login_ip=request.META.get('HTTP_X_FORWARDED_FOR') or request.META.get('REMOTE_ADDR'),
        login_status=False
    )
    logout(request)
    return redirect('backend:backend_login')


# Menu Wise Dashboard
@login_required
def menu_wise_dashboard(request, menu_slug):
    if menu_slug == 'maintenance':
        return redirect('backend:ticket_list')
    current_menu = BackendMenu.objects.filter(menu_url=menu_slug, is_active=True).first()
    if current_menu:
        if current_menu.is_main_menu:
            menu_list = UserMenuPermission.objects.filter(user_id=request.user.id, menu__parent_id=current_menu.id, menu__is_sub_menu=True, can_view=True, menu__is_active=True, is_active=True, deleted=False).order_by('menu__id')
        else:
            menu_list = UserMenuPermission.objects.filter(user_id=request.user.id, menu__parent_id=current_menu.id, menu__is_sub_child_menu=True, can_view=True, menu__is_active=True, is_active=True, deleted=False).order_by('menu__id')
    else:
        menu_list = []

    context = {
        "current_menu": current_menu,
        "menu_list": menu_list,
    }
    return render(request, 'menu_wise_dashboard.html', context)

# Menu Wise Dashboard


# Management Start
@method_decorator(login_required, name='dispatch')
class UserListView(ListView):
    model = User
    template_name = 'user/list.html'
    context_object_name = 'user_list'
    ordering = ['-date_joined']
    paginate_by = 10

    def dispatch(self, request, *args, **kwargs):
        # Check permission before anything else
        if not checkUserPermission(request, "can_add", "/backend/user/"):
            return render(request, "403.html", status=403)
        return super().dispatch(request, *args, **kwargs)

    def get_queryset(self):
        queryset = super().get_queryset()
        username = self.request.GET.get('username')
        is_active = self.request.GET.get('is_active')

        if username:
            queryset = queryset.filter(username__icontains=username)
        if is_active in ['0', '1']:  # stricter check
            queryset = queryset.filter(is_active=is_active)

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['username'] = self.request.GET.get('username', '')
        context['is_active'] = self.request.GET.get('is_active', '')

        get_params = self.request.GET.copy()
        get_params.pop('page', None)
        context['query_params'] = get_params.urlencode()

        context['filter_user'] = User.objects.all()
        return context


@login_required
def user_add(request):
    if not checkUserPermission(request, "can_add", "/backend/user/"):
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        form = UserCreateForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            messages.success(request, 'New user has been added successfully!')
            return redirect('backend:user_list')
        else:
            for field, errors in form.errors.items():
                for error in errors:
                    messages.error(request, f"{field}: {error}")
    else:
        form = UserCreateForm()

    return render(request, 'user/add.html', {'form': form})


@login_required
def user_update(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/user/"):
        return render(request, "403.html", status=403)

    user_obj = get_object_or_404(User, id=data_id)

    if request.method == 'POST':
        # Update logic goes here
        pass

    return render(request, 'user/management_update.html', {"user": user_obj})


@login_required
def reset_password(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/user/"):
        return render(request, "403.html", status=403)

    user = get_object_or_404(User, id=data_id)
    if request.method == 'POST':
        form = AdminPasswordChangeForm(user=user, data=request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Your password has been updated successfully.')
            return redirect('backend:backend_dashboard')
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = AdminPasswordChangeForm(user=user)

    return render(request, 'user/reset_password.html', {'form': form, 'user': user})


@login_required
def user_permission(request, user_id):
    if not checkUserPermission(request, "can_update", "/backend/user/"):
        return render(request, "403.html")

    if request.method == "POST":
        username = request.POST.get("username")
        user_status = request.POST.get("user_status")
        selected_menus = request.POST.getlist("selected_menus")
        can_view = request.POST.getlist("can_view")
        can_add = request.POST.getlist("can_add")
        can_update = request.POST.getlist("can_update")
        can_delete = request.POST.getlist("can_delete")

        try:
            user = User.objects.get(pk=user_id)
            user.is_active = user_status
            user.save()
        except Exception:
            pass

        exist_all_permission = UserMenuPermission.objects.filter(user_id=user_id)
        for exist_permission in exist_all_permission:
            if exist_permission.id not in selected_menus:
                exist_permission.can_view = False
                exist_permission.can_add = False
                exist_permission.can_update = False
                exist_permission.can_delete = False
                exist_permission.is_active = False
                exist_permission.updated_at = datetime.now()
                exist_permission.deleted_by_id = request.user.id
                exist_permission.save()

        if user_id and username and selected_menus:
            for menu_id in selected_menus:
                if menu_id in can_view:
                    user_view_access = True
                else:
                    user_view_access = False

                if menu_id in can_add:
                    user_add_access = True
                else:
                    user_add_access = False

                if menu_id in can_update:
                    user_update_access = True
                else:
                    user_update_access = False

                if menu_id in can_delete:
                    user_delete_access = True
                else:
                    user_delete_access = False

                exist_permission = UserMenuPermission.objects.filter(user_id=user_id, menu_id=menu_id)
                if exist_permission:
                    exist_permission.update(
                        updated_by_id=request.user.id, can_view=user_view_access, can_add=user_add_access,
                        can_update=user_update_access, can_delete=user_delete_access, updated_at=datetime.now(), is_active=True,
                    )
                else:
                    UserMenuPermission.objects.create(
                        user_id=user_id, menu_id=menu_id, can_view=user_view_access, can_add=user_add_access,
                        can_update=user_update_access, can_delete=user_delete_access, created_by_id=request.user.id
                    )
            messages.success(request, "User permission has been assigned!")
        else:
            messages.warning(request, "No permission has been assigned!")

        return redirect('backend:user_permission', user_id=user_id)

    user = User.objects.get(pk=user_id)
    menu_list = BackendMenu.objects.filter(is_active=True).order_by("module_name")

    for data in menu_list:
        try:
            user_access_perm = UserMenuPermission.objects.get(user_id=user_id, menu_id=data.id, is_active=True)

            data.user_menu_id = user_access_perm.menu_id
            data.can_view = user_access_perm.can_view
            data.can_add = user_access_perm.can_add
            data.can_update = user_access_perm.can_update
            data.can_delete = user_access_perm.can_delete
        except Exception:
            pass

    context = {
        "user": user,
        "menu_list": menu_list,
    }
    return render(request, 'user/user_permission.html', context)


#Division
@login_required
def division_list(request):
    if not checkUserPermission(request, "can_view", "/backend/location/division/"):
        return render(request, "403.html", status=403)

    search = request.GET.get('search', '').strip()
    status = request.GET.get('status', '').strip()

    qs = Division.objects.filter(deleted=False).annotate(
        subdivision_count=Count('subdivisions', filter=Q(subdivisions__deleted=False), distinct=True),
        building_count=Count('buildings', filter=Q(buildings__deleted=False), distinct=True),
    ).order_by('id')

    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(code__icontains=search))
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_status': status,
        'status_choices': ['Active', 'Inactive'],
    }
    return render(request, 'division/list.html', context)


@login_required
def division_add(request):
    if not checkUserPermission(request, "can_add", "/backend/location/division/"):
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()
        is_active = request.POST.get('is_active', '1') != '0'

        if not name or not code:
            messages.error(request, 'Name and Code are required.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:division_list')
        elif Division.objects.filter(code=code, deleted=False).exists():
            messages.error(request, 'A division with this code already exists.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:division_list')
        else:
            Division.objects.create(
                name=name, code=code,
                is_active=is_active,
                created_by=request.user,
            )
            messages.success(request, 'Division added successfully.')
            return redirect('backend:division_list')

    return render(request, 'division/add.html', {'action': 'Add'})


@login_required
def division_update(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/location/division/"):
        return render(request, "403.html", status=403)

    obj = get_object_or_404(Division, pk=data_id, deleted=False)

    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()

        if not name or not code:
            messages.error(request, 'Name and Code are required.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:division_list')
        elif Division.objects.filter(code=code, deleted=False).exclude(pk=data_id).exists():
            messages.error(request, 'A division with this code already exists.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:division_list')
        else:
            obj.name = name
            obj.code = code
            if 'is_active' in request.POST:
                obj.is_active = request.POST.get('is_active') != '0'
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Division updated successfully.')
            return redirect('backend:division_list')

    context = {'action': 'Update', 'obj': obj}
    return render(request, 'division/update.html', context)


@login_required
def division_status(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/location/division/"):
        messages.error(request, 'You do not have permission to change status of this division.')
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        obj = get_object_or_404(Division, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Division status changed to {status_text} successfully.')
        return redirect('backend:division_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:division_list')


#SubDivision
@login_required
def subdivision_list(request):
    if not checkUserPermission(request, "can_view", "/backend/location/subdivision/"):
        return render(request, "403.html", status=403)

    search = request.GET.get('search', '').strip()
    division_id = request.GET.get('division', '').strip()
    status = request.GET.get('status', '').strip()

    qs = SubDivision.objects.filter(deleted=False).select_related('division').annotate(
        section_count=Count('sections', filter=Q(sections__deleted=False), distinct=True),
        building_count=Count('buildings', filter=Q(buildings__deleted=False), distinct=True),
    ).order_by('id')

    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(code__icontains=search) | Q(division__name__icontains=search))
    if division_id:
        qs = qs.filter(division_id=division_id)
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_division': division_id,
        'selected_status': status,
        'status_choices': ['Active', 'Inactive'],
        'divisions': Division.objects.filter(deleted=False, is_active=True).order_by('name'),
    }
    return render(request, 'subdivision/list.html', context)


@login_required
def subdivision_add(request):
    if not checkUserPermission(request, "can_add", "/backend/location/subdivision/"):
        return render(request, "403.html", status=403)

    divisions = Division.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        division_id = request.POST.get('division')
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()
        is_active = request.POST.get('is_active', '1') != '0'

        if not division_id or not name or not code:
            messages.error(request, 'Division, Name and Code are required.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:subdivision_list')
        else:
            SubDivision.objects.create(
                division_id=division_id,
                name=name,
                code=code,
                is_active=is_active,
                created_by=request.user,
            )

            messages.success(request, 'Sub-Division added successfully.')
            return redirect('backend:subdivision_list')

    context = {'action': 'Add', 'divisions': divisions}
    return render(request, 'subdivision/add.html', context)


@login_required
def subdivision_update(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/location/subdivision/"):
        return render(request, "403.html", status=403)

    obj = get_object_or_404(SubDivision, pk=data_id, deleted=False)
    divisions = Division.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        division_id = request.POST.get('division')
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()

        if not division_id or not name or not code:
            messages.error(request, 'Division, Name and Code are required.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:subdivision_list')
        else:
            obj.division_id = division_id
            obj.name = name
            obj.code = code
            if 'is_active' in request.POST:
                obj.is_active = request.POST.get('is_active') != '0'
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Sub-Division updated successfully.')
            return redirect('backend:subdivision_list')

    context = {'action': 'Update', 'obj': obj, 'divisions': divisions}
    return render(request, 'subdivision/update.html', context)


@login_required
def subdivision_status(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/location/subdivision/"):
        messages.error(request, 'You do not have permission to change status of this sub-division.')
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        obj = get_object_or_404(SubDivision, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Sub-Division status changed to {status_text} successfully.')
        return redirect('backend:subdivision_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:subdivision_list')


@login_required
def section_list(request):
    if not checkUserPermission(request, "can_view", "/backend/location/section/"):
        return render(request, "403.html", status=403)

    search = request.GET.get('search', '').strip()
    subdivision_id = request.GET.get('subdivision', '').strip()
    status = request.GET.get('status', '').strip()

    qs = Section.objects.filter(deleted=False).select_related('subdivision', 'subdivision__division').annotate(
        building_count=Count('buildings', filter=Q(buildings__deleted=False), distinct=True),
    ).order_by('id')

    if search:
        qs = qs.filter(
            Q(name__icontains=search) |
            Q(code__icontains=search) |
            Q(subdivision__name__icontains=search) |
            Q(subdivision__division__name__icontains=search)
        )
    if subdivision_id:
        qs = qs.filter(subdivision_id=subdivision_id)
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_subdivision': subdivision_id,
        'selected_status': status,
        'status_choices': ['Active', 'Inactive'],
        'subdivisions': SubDivision.objects.filter(deleted=False, is_active=True).select_related('division').order_by('name'),
    }
    return render(request, 'section/list.html', context)


@login_required
def section_add(request):
    if not checkUserPermission(request, "can_add", "/backend/location/section/"):
        return render(request, "403.html", status=403)

    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        subdivision_id = request.POST.get('subdivision')
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()
        is_active = request.POST.get('is_active', '1') != '0'

        if not subdivision_id or not name or not code:
            messages.error(request, 'Sub-Division, Name and Code are required.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:section_list')
        elif Section.objects.filter(subdivision_id=subdivision_id, code=code).exists():
            messages.error(request, 'A section with this code already exists in this sub-division.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:section_list')
        else:
            Section.objects.create(
                subdivision_id=subdivision_id, name=name, code=code,
                is_active=is_active, created_by=request.user,
            )
            messages.success(request, 'Section added successfully.')
            return redirect('backend:section_list')

    context = {'action': 'Add', 'subdivisions': subdivisions}
    return render(request, 'section/add.html', context)


@login_required
def section_update(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/location/section/"):
        return render(request, "403.html", status=403)

    obj = get_object_or_404(Section, pk=data_id, deleted=False)
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        subdivision_id = request.POST.get('subdivision')
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()

        if not subdivision_id or not name or not code:
            messages.error(request, 'Sub-Division, Name and Code are required.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:section_list')
        elif Section.objects.filter(subdivision_id=subdivision_id, code=code).exclude(pk=data_id).exists():
            messages.error(request, 'A section with this code already exists in this sub-division.')
            if request.POST.get('from_modal') == '1':
                return redirect('backend:section_list')
        else:
            obj.subdivision_id = subdivision_id
            obj.name = name
            obj.code = code
            if 'is_active' in request.POST:
                obj.is_active = request.POST.get('is_active') != '0'
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Section updated successfully.')
            return redirect('backend:section_list')

    context = {'action': 'Update', 'obj': obj, 'subdivisions': subdivisions}
    return render(request, 'section/update.html', context)


@login_required
def section_status(request, data_id):
    if not checkUserPermission(request, "can_update", "/backend/location/section/"):
        messages.error(request, 'You do not have permission to change status of this section.')
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        obj = get_object_or_404(Section, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Section status changed to {status_text} successfully.')
        return redirect('backend:section_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:section_list')


# Building
@login_required
def building_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/location/building/'):
        return render(request, '403.html', status=403)

    search = request.GET.get('search', '').strip()
    division_id = request.GET.get('division', '').strip()
    subdivision_id = request.GET.get('subdivision', '').strip()
    section_id = request.GET.get('section', '').strip()
    status = request.GET.get('status', '').strip()

    # KPI stats across all non-deleted buildings
    total_buildings_count = Building.objects.filter(deleted=False).count()
    active_buildings_count = Building.objects.filter(deleted=False, is_active=True).exclude(status__in=['Inactive', 'Under Maintenance', 'Decommissioned']).count()
    if active_buildings_count == 0 and total_buildings_count > 0:
        active_buildings_count = Building.objects.filter(deleted=False, is_active=True).count()
    under_maintenance_count = Building.objects.filter(deleted=False, status='Under Maintenance').count()
    inactive_buildings_count = Building.objects.filter(deleted=False).filter(Q(is_active=False) | Q(status='Inactive')).count()

    qs = Building.objects.filter(deleted=False).select_related('division', 'subdivision', 'section').annotate(
        equipment_count=Count('equipment', filter=Q(equipment__deleted=False))
    ).order_by('building_id', '-created_at')
    
    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(building_id__icontains=search) | Q(surname__icontains=search) | Q(address__icontains=search))
    if division_id:
        qs = qs.filter(division_id=division_id)
    if subdivision_id:
        qs = qs.filter(subdivision_id=subdivision_id)
    if section_id:
        qs = qs.filter(section_id=section_id)
    if status:
        qs = qs.filter(status=status)

    filtered_total_count = qs.count()

    try:
        per_page = int(request.GET.get('per_page', 10))
        if per_page <= 0:
            per_page = 10
    except (ValueError, TypeError):
        per_page = 10

    page_num = request.GET.get('page', 1)
    paginator = Paginator(qs, per_page)
    try:
        data_list = paginator.page(page_num)
    except PageNotAnInteger:
        data_list = paginator.page(1)
    except EmptyPage:
        data_list = paginator.page(paginator.num_pages)

    last_page = paginator.num_pages
    current_page = data_list.number
    start_p = max(1, current_page - 2)
    end_p = min(last_page, current_page + 2)
    paginator_list = list(range(start_p, end_p + 1))
    all_pages = list(range(1, last_page + 1))

    # Fetch filter options
    divisions = Division.objects.filter(deleted=False, is_active=True)
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True)
    sections = Section.objects.filter(deleted=False, is_active=True)
    status_choices = [choice[0] for choice in Building.STATUS_CHOICES]

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'all_pages': all_pages,
        'last_page': last_page,
        'per_page': per_page,
        'total_buildings_kpi': f"{total_buildings_count:02d}",
        'active_buildings_kpi': f"{active_buildings_count:02d}",
        'under_maintenance_kpi': f"{under_maintenance_count:02d}",
        'inactive_buildings_kpi': f"{inactive_buildings_count:02d}",
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_division': division_id,
        'selected_subdivision': subdivision_id,
        'selected_section': section_id,
        'selected_status': status,
        'divisions': divisions,
        'subdivisions': subdivisions,
        'sections': sections,
        'status_choices': status_choices,
    }
    return render(request, 'building/list.html', context)

@login_required
def building_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/location/building/'):
        return render(request, '403.html', status=403)

    divisions = Division.objects.filter(deleted=False, is_active=True).order_by('name')
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True).order_by('name')
    sections = Section.objects.filter(deleted=False, is_active=True).order_by('name')
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')

    if request.method == 'POST':
        building_id = request.POST.get('building_id', '').strip()
        name = request.POST.get('name', '').strip()
        surname = request.POST.get('surname', '').strip()
        year_of_establishment = request.POST.get('year_of_establishment', '').strip() 
        division_id = request.POST.get('division') or None
        subdivision_id = request.POST.get('subdivision') or None
        section_id = request.POST.get('section') or None
        address = request.POST.get('address', '').strip()
        latitude = request.POST.get('latitude')
        longitude = request.POST.get('longitude')
        status = request.POST.get('status', 'Active')

        try:
            latitude = float(latitude) if latitude else None
        except ValueError:
            latitude = None
            
        try:
            longitude = float(longitude) if longitude else None
        except ValueError:
            longitude = None

        has_sections = Section.objects.filter(subdivision__division_id=division_id, deleted=False).exists() if division_id else False

        if not building_id:
            messages.error(request, 'Building ID is required.')
        elif not name:
            messages.error(request, 'Building Name is required.')
        elif has_sections and not section_id:
            messages.error(request, 'Section is required.')
        elif not address:
            messages.error(request, 'Physical Address is required.')
        elif Building.objects.filter(building_id=building_id, deleted=False).exists():
            messages.error(request, 'A building with this Building ID already exists.')
        elif Building.objects.filter(name=name, deleted=False).exists():
            messages.error(request, 'A building with this name already exists.')
        else:
            building = Building.objects.create(
                building_id=building_id,
                name=name,
                surname=surname,
                year_of_establishment=year_of_establishment if year_of_establishment else None, 
                division_id=division_id,
                subdivision_id=subdivision_id,
                section_id=section_id,
                address=address,
                latitude=latitude,
                longitude=longitude,
                status=status,
                created_by=request.user,
            )

            # Process any equipments added in Step 2
            equipments_json = request.POST.get('equipments_data', '').strip()
            if equipments_json:
                try:
                    eq_list = json.loads(equipments_json)
                    for item in eq_list:
                        eq_id = item.get('equipment_id', '').strip()
                        eq_type_id = item.get('equipment_type_id')
                        existing_pk = item.get('id')

                        eq_record = None
                        if existing_pk:
                            eq_record = Equipment.objects.filter(id=existing_pk, deleted=False).first()
                        if not eq_record and eq_id:
                            eq_record = Equipment.objects.filter(equipment_id=eq_id, deleted=False).first()

                        if eq_record:
                            eq_record.building = building
                            if item.get('floor_location'):
                                eq_record.floor_location = item.get('floor_location').strip()
                            eq_record.updated_by = request.user
                            eq_record.save()
                        elif eq_id and eq_type_id:
                            Equipment.objects.create(
                                equipment_id=eq_id,
                                equipment_name=item.get('equipment_name') or item.get('brand', '') + ' ' + (item.get('model_number') or ''),
                                qr_code=eq_id,
                                equipment_type_id=eq_type_id,
                                building=building,
                                brand=item.get('brand', '').strip(),
                                model_number=item.get('model_number', '').strip(),
                                serial_number=item.get('serial_number', '').strip(),
                                floor_location=item.get('floor_location', '').strip(),
                                status='active',
                                maintenance_period_days=int(item.get('maintenance_period_days') or 90),
                                created_by=request.user,
                                is_active=True
                            )
                except Exception as e:
                    pass

            messages.success(request, 'Building added successfully.')
            return redirect('backend:building_list')

    # Fetch all active equipments for the directory picker
    directory_equipments = Equipment.objects.filter(deleted=False, is_active=True).select_related('equipment_type', 'building').order_by('-id')
    directory_equipments_data = [
        {
            'id': eq.id,
            'equipment_id': eq.equipment_id,
            'equipment_name': eq.equipment_name or (f"{eq.brand} {eq.model_number}".strip() if (eq.brand or eq.model_number) else eq.equipment_id),
            'equipment_type_id': str(eq.equipment_type_id) if eq.equipment_type_id else '',
            'equipment_type_name': eq.equipment_type.name if eq.equipment_type else 'Equipment',
            'brand': eq.brand or '',
            'model_number': eq.model_number or '',
            'serial_number': eq.serial_number or '',
            'floor_location': eq.floor_location or '',
            'room_location': eq.room_location or '',
            'status': eq.status or 'active',
            'building_id': eq.building_id or '',
            'building_name': eq.building.name if eq.building else 'Unassigned',
            'is_unassigned': eq.building is None,
            'maintenance_period_days': eq.maintenance_period_days or 90,
        }
        for eq in directory_equipments
    ]

    context = {
        'action': 'Add',
        'divisions': divisions,
        'subdivisions': subdivisions,
        'sections': sections,
        'equipment_types': equipment_types,
        'directory_equipments_json': json.dumps(directory_equipments_data)
    }
    return render(request, 'building/add.html', context)

@login_required
def building_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/location/building/'):
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Building, pk=data_id, deleted=False)
    divisions = Division.objects.filter(deleted=False, is_active=True).order_by('name')
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True).order_by('name')
    sections = Section.objects.filter(deleted=False, is_active=True).order_by('name')
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')

    # Fetch existing equipments for this building
    existing_equipments = Equipment.objects.filter(building=obj, deleted=False).select_related('equipment_type')
    existing_equipments_data = [
        {
            'id': eq.id,
            'equipment_id': eq.equipment_id,
            'equipment_name': eq.equipment_name or (f"{eq.brand} {eq.model_number}".strip() if (eq.brand or eq.model_number) else eq.equipment_id),
            'equipment_type_id': str(eq.equipment_type_id) if eq.equipment_type_id else '',
            'equipment_type_name': eq.equipment_type.name if eq.equipment_type else '',
            'brand': eq.brand or '',
            'floor_location': eq.floor_location or '',
            'model_number': eq.model_number or '',
            'serial_number': eq.serial_number or '',
            'status': eq.status or 'active',
            'maintenance_period_days': eq.maintenance_period_days or 90,
        }
        for eq in existing_equipments
    ]

    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        surname = request.POST.get('surname', '').strip()
        year_of_establishment = request.POST.get('year_of_establishment', '').strip()
        division_id = request.POST.get('division') or None
        subdivision_id = request.POST.get('subdivision') or None
        section_id = request.POST.get('section') or None
        address = request.POST.get('address', '').strip()
        latitude = request.POST.get('latitude')
        longitude = request.POST.get('longitude')
        status = request.POST.get('status', obj.status or 'Active')

        try:
            latitude = float(latitude) if latitude else None
        except ValueError:
            latitude = None
            
        try:
            longitude = float(longitude) if longitude else None
        except ValueError:
            longitude = None

        has_sections = Section.objects.filter(subdivision__division_id=division_id, deleted=False).exists() if division_id else False

        if not name:
            messages.error(request, 'Building Name is required.')
        elif has_sections and not section_id:
            messages.error(request, 'Section is required.')
        elif not address:
            messages.error(request, 'Physical Address is required.')
        elif Building.objects.filter(name=name, deleted=False).exclude(pk=data_id).exists():
            messages.error(request, 'A building with this name already exists.')
        else:
            obj.name = name
            obj.surname = surname
            obj.year_of_establishment = year_of_establishment if year_of_establishment else None
            obj.division_id = division_id
            obj.subdivision_id = subdivision_id
            obj.section_id = section_id
            obj.address = address
            obj.latitude = latitude
            obj.longitude = longitude
            obj.status = status
            obj.updated_by = request.user
            obj.save()

            # Process any equipments updated/added in Step 2
            equipments_json = request.POST.get('equipments_data', '').strip()
            if equipments_json:
                try:
                    eq_list = json.loads(equipments_json)
                    submitted_ids = set()
                    for item in eq_list:
                        eq_id = item.get('equipment_id', '').strip()
                        eq_type_id = item.get('equipment_type_id') or None
                        existing_pk = item.get('id')

                        eq_record = None
                        if existing_pk:
                            eq_record = Equipment.objects.filter(id=existing_pk, deleted=False).first()
                        if not eq_record and eq_id:
                            eq_record = Equipment.objects.filter(equipment_id=eq_id, deleted=False).first()

                        if eq_record:
                            submitted_ids.add(eq_record.equipment_id)
                            eq_record.building = obj
                            if eq_type_id:
                                eq_record.equipment_type_id = eq_type_id
                            if item.get('floor_location'):
                                eq_record.floor_location = item.get('floor_location').strip()
                            if item.get('brand'):
                                eq_record.brand = item.get('brand').strip()
                            if item.get('model_number'):
                                eq_record.model_number = item.get('model_number').strip()
                            if item.get('serial_number'):
                                eq_record.serial_number = item.get('serial_number').strip()
                            if item.get('maintenance_period_days'):
                                try:
                                    eq_record.maintenance_period_days = int(item.get('maintenance_period_days'))
                                except (ValueError, TypeError):
                                    pass
                            eq_record.updated_by = request.user
                            eq_record.save()
                        elif eq_id and eq_type_id:
                            submitted_ids.add(eq_id)
                            Equipment.objects.create(
                                equipment_id=eq_id,
                                equipment_name=item.get('equipment_name') or item.get('brand', '') + ' ' + (item.get('model_number') or ''),
                                qr_code=eq_id,
                                equipment_type_id=eq_type_id,
                                building=obj,
                                brand=item.get('brand', '').strip(),
                                floor_location=item.get('floor_location', '').strip(),
                                model_number=item.get('model_number', '').strip(),
                                serial_number=item.get('serial_number', '').strip(),
                                maintenance_period_days=int(item.get('maintenance_period_days') or 90),
                                is_active=True,
                                deleted=False,
                                created_by=request.user
                            )
                    # Unassign equipments that belonged to this building but were removed
                    Equipment.objects.filter(building=obj, deleted=False).exclude(equipment_id__in=submitted_ids).update(building=None, updated_by=request.user)
                except Exception as e:
                    print('Error updating equipments in building_update:', e)

            messages.success(request, 'Building updated successfully.')
            return redirect('backend:building_list')

    # Directory of all active equipments for picker
    directory_equipments = Equipment.objects.filter(deleted=False, is_active=True).select_related('equipment_type', 'building').order_by('-id')
    directory_equipments_data = [
        {
            'id': eq.id,
            'equipment_id': eq.equipment_id,
            'equipment_name': eq.equipment_name or (f"{eq.brand} {eq.model_number}".strip() if (eq.brand or eq.model_number) else eq.equipment_id),
            'equipment_type_id': str(eq.equipment_type_id) if eq.equipment_type_id else '',
            'equipment_type_name': eq.equipment_type.name if eq.equipment_type else 'Equipment',
            'brand': eq.brand or '',
            'model_number': eq.model_number or '',
            'serial_number': eq.serial_number or '',
            'floor_location': eq.floor_location or '',
            'room_location': eq.room_location or '',
            'status': eq.status or 'active',
            'building_id': eq.building_id or '',
            'building_name': eq.building.name if eq.building else 'Unassigned',
            'is_unassigned': eq.building is None,
            'maintenance_period_days': eq.maintenance_period_days or 90,
        }
        for eq in directory_equipments
    ]

    context = {
        'action': 'Update',
        'obj': obj,
        'divisions': divisions,
        'subdivisions': subdivisions,
        'sections': sections,
        'equipment_types': equipment_types,
        'existing_equipments_json': json.dumps(existing_equipments_data),
        'directory_equipments_json': json.dumps(directory_equipments_data)
    }
    return render(request, 'building/update.html', context)

@login_required
def building_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/location/building/'):
        messages.error(request, 'You do not have permission to change status of this building.')
        return render(request, '403.html', status=403)

    if request.method == 'POST':
        obj = get_object_or_404(Building, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Building status changed to {status_text} successfully.')
        return redirect('backend:building_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:building_list')

@login_required
def building_details(request, data_id):
    if not checkUserPermission(request, 'can_view', '/backend/location/building/'):
        return render(request, '403.html', status=403)

    building = get_object_or_404(
        Building.objects.select_related('division', 'subdivision', 'section'),
        pk=data_id,
        deleted=False,
    )

    # Handle assigning unassigned equipments directly from the Add Equipment modal
    if request.method == 'POST' and request.POST.get('action') == 'assign_equipments':
        if not checkUserPermission(request, 'can_edit', '/backend/location/building/'):
            messages.error(request, 'You do not have permission to assign equipment.')
            return render(request, '403.html', status=403)

        payload_raw = (request.POST.get('equipments_payload') or '').strip()
        raw_ids = (request.POST.get('equipment_ids') or '').strip()
        items = []
        if payload_raw:
            try:
                parsed = json.loads(payload_raw)
                if isinstance(parsed, list):
                    items = parsed
            except Exception:
                items = []
        if not items and raw_ids:
            items = [{'id': int(x)} for x in raw_ids.split(',') if x.strip().isdigit()]

        if items:
            updated_count = 0
            for item in items:
                try:
                    eq_pk = int(item.get('id'))
                except (TypeError, ValueError):
                    continue
                eq = Equipment.objects.filter(pk=eq_pk, deleted=False, building__isnull=True).first()
                if not eq:
                    continue

                eq.building = building
                if 'floor_location' in item:
                    eq.floor_location = (item.get('floor_location') or '').strip()
                if 'room_location' in item:
                    eq.room_location = (item.get('room_location') or '').strip()
                if 'installation_notes' in item:
                    eq.installation_notes = (item.get('installation_notes') or '').strip()
                if 'maintenance_notes' in item:
                    eq.maintenance_notes = (item.get('maintenance_notes') or '').strip()

                inst_str = (item.get('installment_date') or '').strip()
                if inst_str:
                    try:
                        eq.installment_date = timezone.make_aware(datetime.strptime(inst_str, '%Y-%m-%d'))
                    except ValueError:
                        pass

                freq_unit = (item.get('frequency_unit') or eq.frequency_unit or 'Days').strip() or 'Days'
                eq.frequency_unit = freq_unit
                try:
                    mpd = int(item.get('maintenance_period_days') or eq.maintenance_period_days or 90)
                    eq.maintenance_period_days = mpd if mpd > 0 else 90
                except (TypeError, ValueError):
                    pass

                last_str = (item.get('last_maintenance_date') or '').strip()
                if last_str:
                    try:
                        eq.last_maintenance_date = datetime.strptime(last_str, '%Y-%m-%d').date()
                    except ValueError:
                        pass

                next_str = (item.get('next_service_due') or '').strip()
                if next_str:
                    try:
                        eq.next_service_due = datetime.strptime(next_str, '%Y-%m-%d').date()
                    except ValueError:
                        pass
                elif eq.last_maintenance_date:
                    if eq.frequency_unit == 'Months':
                        eq.next_service_due = eq.last_maintenance_date + timedelta(days=eq.maintenance_period_days * 30)
                    elif eq.frequency_unit == 'Years':
                        eq.next_service_due = eq.last_maintenance_date + timedelta(days=eq.maintenance_period_days * 365)
                    else:
                        eq.next_service_due = eq.last_maintenance_date + timedelta(days=eq.maintenance_period_days)

                eq.updated_by = request.user
                eq.save()
                updated_count += 1

            if updated_count:
                messages.success(request, f'Assigned {updated_count} equipment to {building.name}.')
            else:
                messages.info(request, 'No unassigned equipment was updated.')
        else:
            messages.warning(request, 'Please select at least one equipment to add.')
        return redirect('backend:building_details', data_id=building.id)

    today = date.today()
    due_soon_cutoff = today + timedelta(days=30)

    # All equipment in this building (unfiltered) for overview & maintenance status cards
    all_equipment = Equipment.objects.filter(
        building=building,
        deleted=False,
    ).select_related('equipment_type').order_by('id')

    total_assets = all_equipment.count()
    active_count = all_equipment.filter(status='active').count()
    maintenance_count = all_equipment.filter(status='under_maintenance').count()
    in_repair_count = all_equipment.filter(status='in_repair').count()
    inactive_count = all_equipment.filter(status__in=['inactive', 'decommissioned']).count()
    equipment_types_count = (
        all_equipment.values('equipment_type_id').distinct().count()
    )

    # Maintenance status counts for this building's equipment
    scheduled_count = all_equipment.filter(next_service_due__isnull=False, next_service_due__gt=due_soon_cutoff).count()
    overdue_count = all_equipment.filter(next_service_due__isnull=False, next_service_due__lt=today).count()
    due_soon_count = all_equipment.filter(
        next_service_due__isnull=False,
        next_service_due__gte=today,
        next_service_due__lte=due_soon_cutoff,
    ).count()

    # Filter options for table
    filter_types = EquipmentType.objects.filter(
        id__in=all_equipment.values_list('equipment_type_id', flat=True).distinct()
    ).order_by('name')
    filter_brands = list(
        all_equipment.exclude(brand='').exclude(brand__isnull=True)
        .values_list('brand', flat=True).distinct().order_by('brand')
    )

    # Apply GET filters to table queryset
    selected_type = (request.GET.get('type') or '').strip()
    selected_brand = (request.GET.get('brand') or '').strip()
    selected_status = (request.GET.get('status') or '').strip()
    selected_maint = (request.GET.get('maint') or '').strip()

    equipment_list = all_equipment
    if selected_type:
        equipment_list = equipment_list.filter(equipment_type_id=selected_type)
    if selected_brand:
        equipment_list = equipment_list.filter(brand=selected_brand)
    if selected_status:
        equipment_list = equipment_list.filter(status=selected_status)
    if selected_maint == 'overdue':
        equipment_list = equipment_list.filter(next_service_due__isnull=False, next_service_due__lt=today)
    elif selected_maint == 'due_soon':
        equipment_list = equipment_list.filter(
            next_service_due__isnull=False,
            next_service_due__gte=today,
            next_service_due__lte=due_soon_cutoff,
        )
    elif selected_maint == 'scheduled':
        equipment_list = equipment_list.filter(next_service_due__isnull=False, next_service_due__gt=due_soon_cutoff)

    # Pagination (matching includes/pagination.html)
    try:
        per_page = int(request.GET.get('per_page', 10))
        if per_page not in (10, 20, 50, 100):
            per_page = 10
    except (TypeError, ValueError):
        per_page = 10

    filtered_count = equipment_list.count()
    paginator = Paginator(equipment_list, per_page)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)

    # Unassigned equipment directory for the Add Equipment popup modal
    modal_equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')
    unassigned_qs = Equipment.objects.filter(
        deleted=False,
        building__isnull=True,
    ).select_related('equipment_type').order_by('-id')
    unassigned_equipments_data = [
        {
            'id': eq.id,
            'equipment_id': eq.equipment_id or f'EQ-{eq.id}',
            'equipment_name': eq.equipment_name or (eq.equipment_type.name if eq.equipment_type else 'Equipment'),
            'equipment_type_id': eq.equipment_type_id,
            'equipment_type_name': eq.equipment_type.name if eq.equipment_type else 'Equipment',
            'brand': eq.brand or '',
            'model_number': eq.model_number or '',
            'serial_number': eq.serial_number or '',
            'floor_location': eq.floor_location or '',
            'room_location': eq.room_location or '',
            'installment_date': eq.installment_date.strftime('%Y-%m-%d') if eq.installment_date else today.strftime('%Y-%m-%d'),
            'installation_notes': eq.installation_notes or '',
            'maintenance_period_days': eq.maintenance_period_days or 90,
            'frequency_unit': eq.frequency_unit or 'Days',
            'last_maintenance_date': eq.last_maintenance_date.strftime('%Y-%m-%d') if eq.last_maintenance_date else '',
            'next_service_due': eq.next_service_due.strftime('%Y-%m-%d') if eq.next_service_due else '',
            'maintenance_notes': eq.maintenance_notes or '',
            'is_unassigned': True,
        }
        for eq in unassigned_qs
    ]

    context = {
        'building': building,
        'equipment_list': page_obj,
        'page_obj': page_obj,
        'per_page': per_page,
        'filtered_count': filtered_count,
        'total_assets': total_assets,
        'active_count': active_count,
        'maintenance_count': maintenance_count,
        'in_repair_count': in_repair_count,
        'inactive_count': inactive_count,
        'equipment_types_count': equipment_types_count,
        'scheduled_count': scheduled_count,
        'overdue_count': overdue_count,
        'due_soon_count': due_soon_count,
        'filter_types': filter_types,
        'filter_brands': filter_brands,
        'equipment_types': modal_equipment_types,
        'directory_equipments_json': json.dumps(unassigned_equipments_data),
        'status_choices': Equipment.STATUS_CHOICES,
        'selected_type': selected_type,
        'selected_brand': selected_brand,
        'selected_status': selected_status,
        'selected_maint': selected_maint,
        'today': today,
    }
    return render(request, 'building/details.html', context)


@login_required
def dash_board(request):
    today = date.today()

    # --- Parse Filters ---
    division_id = (request.GET.get('division') or '').strip()
    subdivision_id = (request.GET.get('subdivision') or '').strip()
    section_id = (request.GET.get('section') or '').strip()
    start_date_param = (request.GET.get('start_date') or '').strip()
    end_date_param = (request.GET.get('end_date') or '').strip()

    try:
        end_date = datetime.strptime(end_date_param, '%Y-%m-%d').date() if end_date_param else today
    except ValueError:
        end_date = today

    try:
        start_date = datetime.strptime(start_date_param, '%Y-%m-%d').date() if start_date_param else (end_date - timedelta(days=16))
    except ValueError:
        start_date = end_date - timedelta(days=16)

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    span_days = max(1, (end_date - start_date).days + 1)
    prev_end_date = start_date - timedelta(days=1)
    prev_start_date = prev_end_date - timedelta(days=span_days - 1)

    if start_date.month == end_date.month and start_date.year == end_date.year:
        date_range_label = f"{start_date.strftime('%b %d')} to {end_date.strftime('%d')}"
    else:
        date_range_label = f"{start_date.strftime('%b %d')} to {end_date.strftime('%b %d')}"

    if prev_start_date.month == prev_end_date.month and prev_start_date.year == prev_end_date.year:
        prev_range_label = f"vs {prev_start_date.strftime('%b %d')} to {prev_end_date.strftime('%d')}"
    else:
        prev_range_label = f"vs {prev_start_date.strftime('%b %d')} to {prev_end_date.strftime('%b %d')}"

    # --- Dropdown Options ---
    divisions = Division.objects.filter(deleted=False, is_active=True).order_by('name')
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True).order_by('name')
    if division_id.isdigit():
        subdivisions = subdivisions.filter(division_id=int(division_id))
    sections = Section.objects.filter(deleted=False, is_active=True).order_by('name')
    if subdivision_id.isdigit():
        sections = sections.filter(subdivision_id=int(subdivision_id))
    elif division_id.isdigit():
        sections = sections.filter(subdivision__division_id=int(division_id))

    # --- Building stats ---
    buildings_qs = Building.objects.filter(deleted=False)
    if division_id.isdigit():
        buildings_qs = buildings_qs.filter(division_id=int(division_id))
    if subdivision_id.isdigit():
        buildings_qs = buildings_qs.filter(subdivision_id=int(subdivision_id))
    if section_id.isdigit():
        buildings_qs = buildings_qs.filter(section_id=int(section_id))

    total_buildings = buildings_qs.count()
    active_buildings = buildings_qs.filter(status='Active').count()
    inactive_buildings = buildings_qs.filter(status='Inactive').count()
    renovation_buildings = buildings_qs.filter(status='Under Maintenance').count()

    # --- Equipment stats ---
    equipment_qs = Equipment.objects.filter(deleted=False, building__in=buildings_qs)
    total_equipment = equipment_qs.count()
    eq_active = equipment_qs.filter(status='active').count()
    eq_maintenance = equipment_qs.filter(status='under_maintenance').count()
    eq_repair = equipment_qs.filter(status='in_repair').count()
    eq_outdated = equipment_qs.filter(status='inactive').count()
    eq_decommissioned = equipment_qs.filter(status='decommissioned').count()

    if total_equipment > 0:
        uptime_pct = round(eq_active / total_equipment * 100, 1)
        repair_pct = round(eq_repair / total_equipment * 100, 1)
        outdated_pct = round(eq_outdated / total_equipment * 100, 1)
    else:
        uptime_pct = repair_pct = outdated_pct = 0

    # --- Tickets ---
    tickets_qs = Ticket.objects.filter(deleted=False, building__in=buildings_qs)
    open_tickets = tickets_qs.filter(status__in=['open', 'assigned', 'in_progress'])
    open_ticket_count = open_tickets.count()
    critical_ticket_count = open_tickets.filter(priority='critical').count()
    assigned_ticket_count = open_tickets.filter(status__in=['assigned', 'in_progress']).count()

    unassigned_tickets_qs = (
        tickets_qs.filter(status='open')
        .select_related('equipment', 'building', 'equipment__equipment_type')
        .order_by('-opened_at')
    )
    unassigned_ticket_count = unassigned_tickets_qs.count()
    unassigned_tickets = list(unassigned_tickets_qs[:3])
    recent_tickets = (
        open_tickets
        .select_related('equipment', 'building', 'equipment__equipment_type')
        .order_by('-opened_at')[:5]
    )

    # --- Maintenance Records & Scheduled Maintenance ---
    maint_records_qs = MaintenanceRecord.objects.filter(
        deleted=False,
        equipment__in=equipment_qs,
    ).select_related('equipment', 'equipment__equipment_type', 'equipment__building', 'assigned_to', 'assigned_to__user', 'technician', 'maintenance')

    sched_maint_qs = ScheduledMaintenance.objects.filter(
        deleted=False,
        equipment__in=equipment_qs,
    ).select_related('equipment', 'equipment__equipment_type', 'equipment__building', 'assigned_to')

    active_maintenance_count = (
        maint_records_qs.filter(maintainance_status__in=['pending', 'in_progress']).count()
        + sched_maint_qs.filter(status='pending', scheduled_date__gte=today).count()
    )

    # --- Overdue Maintenance (Scheduled + Pending Records past due) ---
    overdue_sched = list(sched_maint_qs.filter(status='pending', scheduled_date__lt=today).order_by('scheduled_date'))
    overdue_records = list(
        maint_records_qs.filter(maintainance_status__in=['pending', 'in_progress'], maintenance_date__lt=today)
        .order_by('maintenance_date')
    )

    overdue_items = []
    for sm in overdue_sched:
        days_late = max(1, (today - sm.scheduled_date).days)
        overdue_items.append({
            'title': sm.title or (sm.equipment.equipment_type.name if sm.equipment and sm.equipment.equipment_type else 'Scheduled Maintenance'),
            'days_overdue': days_late,
            'equipment_id': sm.equipment.equipment_id if sm.equipment else '—',
            'building_name': sm.equipment.building.name if sm.equipment and sm.equipment.building else '—',
            'scheduled_date': sm.scheduled_date,
            'detail_url': reverse('backend:equipment_detail', args=[sm.equipment.id]) if sm.equipment else reverse('backend:maintenance'),
        })
    for mr in overdue_records:
        days_late = max(1, (today - mr.maintenance_date).days)
        eq_label = mr.equipment.equipment_type.name if mr.equipment and mr.equipment.equipment_type else 'Maintenance'
        overdue_items.append({
            'title': f"{eq_label} ({mr.get_record_type_display()})",
            'days_overdue': days_late,
            'equipment_id': mr.equipment.equipment_id if mr.equipment else '—',
            'building_name': mr.equipment.building.name if mr.equipment and mr.equipment.building else '—',
            'scheduled_date': mr.maintenance_date,
            'detail_url': reverse('backend:maintenance_record_detail', args=[mr.id]),
        })
    overdue_items.sort(key=lambda x: x['scheduled_date'])
    overdue_count = len(overdue_items)
    overdue_preview = overdue_items[:3]

    # --- Upcoming Maintenance Table (includes overdue + upcoming pending/in_progress) ---
    upcoming_rows = []
    for sm in sched_maint_qs.filter(status='pending').order_by('scheduled_date')[:10]:
        is_overdue = sm.scheduled_date < today
        tech_name = (sm.assigned_to.get_full_name() or sm.assigned_to.username) if sm.assigned_to else 'Unassigned'
        upcoming_rows.append({
            'equipment_name': sm.title or (sm.equipment.equipment_type.name if sm.equipment and sm.equipment.equipment_type else 'Equipment'),
            'equipment_pk': sm.equipment.id if sm.equipment else None,
            'equipment_code': sm.equipment.equipment_id if sm.equipment else '—',
            'building_name': sm.equipment.building.name if sm.equipment and sm.equipment.building else '—',
            'building_pk': sm.equipment.building.id if sm.equipment and sm.equipment.building else None,
            'maint_type': sm.get_recurrence_display() if sm.recurrence != 'once' else 'Routine',
            'due_date': sm.scheduled_date,
            'technician_name': tech_name,
            'status_label': 'Overdue' if is_overdue else 'Pending',
            'status_key': 'overdue' if is_overdue else 'pending',
            'action_url': reverse('backend:equipment_detail', args=[sm.equipment.id]) if sm.equipment else reverse('backend:maintenance'),
        })

    for mr in maint_records_qs.filter(maintainance_status__in=['pending', 'in_progress']).order_by('maintenance_date')[:10]:
        is_overdue = mr.maintenance_date < today
        if mr.assigned_to:
            tech_name = mr.assigned_to.display_name
        elif mr.technician:
            tech_name = mr.technician.get_full_name() or mr.technician.username
        else:
            tech_name = 'Unassigned'
        status_key = 'overdue' if is_overdue else mr.maintainance_status
        status_label = 'Overdue' if is_overdue else mr.get_maintainance_status_display()
        upcoming_rows.append({
            'equipment_name': mr.equipment.equipment_type.name if mr.equipment and mr.equipment.equipment_type else 'Equipment',
            'equipment_pk': mr.equipment.id if mr.equipment else None,
            'equipment_code': mr.equipment.equipment_id if mr.equipment else '—',
            'building_name': mr.equipment.building.name if mr.equipment and mr.equipment.building else '—',
            'building_pk': mr.equipment.building.id if mr.equipment and mr.equipment.building else None,
            'maint_type': mr.get_record_type_display(),
            'due_date': mr.maintenance_date,
            'technician_name': tech_name,
            'status_label': status_label,
            'status_key': status_key,
            'action_url': reverse('backend:maintenance_record_detail', args=[mr.id]),
        })

    upcoming_rows.sort(key=lambda x: x['due_date'])
    upcoming_maintenance_count = len(upcoming_rows)
    upcoming_table_rows = upcoming_rows[:10]

    # --- Critical Maintenance Requests ---
    critical_requests = []
    crit_tickets = (
        open_tickets.filter(priority='critical')
        .select_related('equipment', 'equipment__equipment_type', 'building')
        .order_by('-opened_at')
    )
    for t in crit_tickets[:3]:
        eq_name = t.equipment.equipment_type.name if t.equipment and t.equipment.equipment_type else t.title
        bldg_name = t.building.name if t.building else '—'
        desc = t.issue_description or t.title
        critical_requests.append({
            'code': f"MR-{t.opened_at.strftime('%Y')}-{t.id:03d}",
            'heading': f"{eq_name} · {bldg_name}",
            'subtitle': desc,
            'time_obj': t.opened_at,
            'detail_url': reverse('backend:ticket_detail', args=[t.id]),
        })

    alerts_qs = CriticalAlert.objects.filter(is_acknowledged=False).order_by('-id')
    if division_id.isdigit() or subdivision_id.isdigit() or section_id.isdigit():
        alerts_qs = alerts_qs.filter(Q(building__in=buildings_qs) | Q(equipment__in=equipment_qs))
    new_alert_count = alerts_qs.count()
    critical_alerts = alerts_qs[:5]

    if len(critical_requests) < 3:
        for alert in alerts_qs[: (3 - len(critical_requests))]:
            bldg_name = alert.building.name if alert.building else 'General'
            critical_requests.append({
                'code': f"AL-{today.year}-{alert.id:03d}",
                'heading': f"{alert.title} · {bldg_name}",
                'subtitle': alert.description or alert.get_alert_type_display(),
                'time_obj': None,
                'detail_url': reverse('backend:equipment_detail', args=[alert.equipment.id]) if alert.equipment else reverse('backend:ticket_list'),
            })
    critical_requests_count = crit_tickets.count() + new_alert_count

    # --- Maintenance Activity Trend (7, 30, 90 days) ---
    def _build_trend(days_window):
        bucket_count = 5
        window_end = end_date
        window_start = window_end - timedelta(days=days_window - 1)
        prev_win_end = window_start - timedelta(days=1)
        prev_win_start = prev_win_end - timedelta(days=days_window - 1)

        req_tickets = tickets_qs.filter(opened_at__date__gte=window_start, opened_at__date__lte=window_end)
        req_records = maint_records_qs.filter(created_at__date__gte=window_start, created_at__date__lte=window_end)
        comp_tickets = tickets_qs.filter(status__in=['resolved', 'closed'], updated_at__date__gte=window_start, updated_at__date__lte=window_end)
        comp_records = maint_records_qs.filter(maintainance_status='completed', updated_at__date__gte=window_start, updated_at__date__lte=window_end)

        prev_req = (
            tickets_qs.filter(opened_at__date__gte=prev_win_start, opened_at__date__lte=prev_win_end).count()
            + maint_records_qs.filter(created_at__date__gte=prev_win_start, created_at__date__lte=prev_win_end).count()
        )
        prev_comp = (
            tickets_qs.filter(status__in=['resolved', 'closed'], updated_at__date__gte=prev_win_start, updated_at__date__lte=prev_win_end).count()
            + maint_records_qs.filter(maintainance_status='completed', updated_at__date__gte=prev_win_start, updated_at__date__lte=prev_win_end).count()
        )

        labels = []
        requested_series = []
        completed_series = []
        step = max(1, days_window // bucket_count)

        for i in range(bucket_count):
            b_start = window_start + timedelta(days=i * step)
            b_end = window_end if i == bucket_count - 1 else min(window_end, b_start + timedelta(days=step - 1))
            labels.append(b_start.strftime('%b %d'))
            r_cnt = (
                req_tickets.filter(opened_at__date__gte=b_start, opened_at__date__lte=b_end).count()
                + req_records.filter(created_at__date__gte=b_start, created_at__date__lte=b_end).count()
            )
            c_cnt = (
                comp_tickets.filter(updated_at__date__gte=b_start, updated_at__date__lte=b_end).count()
                + comp_records.filter(updated_at__date__gte=b_start, updated_at__date__lte=b_end).count()
            )
            requested_series.append(r_cnt)
            completed_series.append(c_cnt)

        tot_req = sum(requested_series)
        tot_comp = sum(completed_series)
        comp_rate = round((tot_comp / tot_req) * 100) if tot_req > 0 else 0
        comp_rate = min(100, comp_rate)

        req_delta = round(((tot_req - prev_req) / prev_req) * 100) if prev_req > 0 else (100 if tot_req > 0 else 0)
        comp_delta = round(((tot_comp - prev_comp) / prev_comp) * 100) if prev_comp > 0 else (100 if tot_comp > 0 else 0)

        return {
            'labels': labels,
            'requested': requested_series,
            'completed': completed_series,
            'total_requested': tot_req,
            'total_completed': tot_comp,
            'requested_delta': req_delta,
            'completed_delta': comp_delta,
            'completion_rate': comp_rate,
        }

    trend_7 = _build_trend(7)
    trend_30 = _build_trend(30)
    trend_90 = _build_trend(90)
    trend_data_json = json.dumps({
        '7': trend_7,
        '30': trend_30,
        '90': trend_90,
    })

    # --- Recent Activities Timeline ---
    recent_activities = []
    ticket_logs = (
        TicketActivityLog.objects.filter(ticket__in=tickets_qs)
        .select_related('ticket', 'actor', 'ticket__building')
        .order_by('-created_at')[:5]
    )
    for log in ticket_logs:
        actor_name = (log.actor.get_full_name() or log.actor.username) if log.actor else 'System'
        initials = ''.join([p[0] for p in actor_name.split()[:2]]).upper() or 'SY'
        ref_code = f"WO-{log.ticket.opened_at.strftime('%Y')}-{log.ticket.id:03d}" if log.ticket else ''
        recent_activities.append({
            'title': log.get_action_type_display(),
            'badge_label': log.get_action_type_display().upper(),
            'actor_name': actor_name,
            'actor_initials': initials,
            'description': log.description or f"Activity recorded on {ref_code}",
            'timestamp': log.created_at,
            'time_str': format_local_time(log.created_at, '%I:%M %p'),
            'ref_code': ref_code,
            'ref_url': reverse('backend:ticket_detail', args=[log.ticket.id]) if log.ticket else reverse('backend:ticket_list'),
        })

    if len(recent_activities) < 5:
        for mr in maint_records_qs.order_by('-updated_at')[: (5 - len(recent_activities))]:
            actor_obj = mr.updated_by or mr.created_by or mr.technician
            if mr.assigned_to and not actor_obj:
                actor_name = mr.assigned_to.display_name
            elif actor_obj:
                actor_name = actor_obj.get_full_name() or actor_obj.username
            else:
                actor_name = 'System'
            initials = ''.join([p[0] for p in actor_name.split()[:2]]).upper() or 'SY'
            bldg_name = mr.equipment.building.name if mr.equipment and mr.equipment.building else 'Facility'
            eq_name = mr.equipment.equipment_type.name if mr.equipment and mr.equipment.equipment_type else 'Equipment'
            status_title = f"Maintenance {mr.get_maintainance_status_display()}"
            ref_code = mr.maintenance.maintenance_serial if mr.maintenance else f"MR-{mr.maintenance_date.year}-{mr.id:03d}"
            recent_activities.append({
                'title': status_title,
                'badge_label': status_title.upper(),
                'actor_name': actor_name,
                'actor_initials': initials,
                'description': mr.work_description or f"{mr.get_record_type_display()} maintenance on {eq_name} at {bldg_name}.",
                'timestamp': mr.updated_at,
                'time_str': format_local_time(mr.updated_at, '%I:%M %p'),
                'ref_code': ref_code,
                'ref_url': reverse('backend:maintenance_record_detail', args=[mr.id]),
            })

    recent_activities.sort(key=lambda x: x['timestamp'], reverse=True)
    recent_activities = recent_activities[:5]
    recent_activities_count = len(recent_activities)

    context = {
        # Filter state
        'divisions': divisions,
        'subdivisions': subdivisions,
        'sections': sections,
        'selected_division': division_id,
        'selected_subdivision': subdivision_id,
        'selected_section': section_id,
        'start_date_str': start_date.strftime('%Y-%m-%d'),
        'end_date_str': end_date.strftime('%Y-%m-%d'),
        'date_range_label': date_range_label,
        'prev_range_label': prev_range_label,
        # Buildings
        'total_buildings': total_buildings,
        'total_buildings_padded': f"{total_buildings:02d}",
        'active_buildings': active_buildings,
        'inactive_buildings': inactive_buildings,
        'renovation_buildings': renovation_buildings,
        # Equipment
        'total_equipment': total_equipment,
        'total_equipment_padded': f"{total_equipment:02d}",
        'eq_active': eq_active,
        'eq_active_padded': f"{eq_active:02d}",
        'eq_maintenance': eq_maintenance,
        'eq_maintenance_padded': f"{eq_maintenance:02d}",
        'eq_repair': eq_repair,
        'eq_repair_padded': f"{eq_repair:02d}",
        'eq_outdated': eq_outdated,
        'eq_inactive_padded': f"{eq_outdated:02d}",
        'eq_decommissioned': eq_decommissioned,
        'eq_decommissioned_padded': f"{eq_decommissioned:02d}",
        'uptime_pct': uptime_pct,
        'repair_pct': repair_pct,
        'outdated_pct': outdated_pct,
        # Tickets
        'open_ticket_count': open_ticket_count,
        'critical_ticket_count': critical_ticket_count,
        'assigned_ticket_count': assigned_ticket_count,
        'unassigned_ticket_count': unassigned_ticket_count,
        'unassigned_tickets': unassigned_tickets,
        'recent_tickets': recent_tickets,
        # Maintenance & Overdue
        'active_maintenance_count': active_maintenance_count,
        'active_maintenance_padded': f"{active_maintenance_count:02d}",
        'upcoming_maintenance_count': upcoming_maintenance_count,
        'upcoming_table_rows': upcoming_table_rows,
        'overdue_count': overdue_count,
        'overdue_count_padded': f"{overdue_count:02d}",
        'overdue_preview': overdue_preview,
        # Critical Requests & Alerts
        'new_alert_count': new_alert_count,
        'critical_alerts': critical_alerts,
        'critical_requests': critical_requests,
        'critical_requests_count': critical_requests_count,
        # Trend Chart
        'trend_7': trend_7,
        'trend_data_json': trend_data_json,
        # Recent Activities
        'recent_activities': recent_activities,
        'recent_activities_count': recent_activities_count,
    }
    return render(request, 'dashboard.html', context)




@login_required
def technician_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/technician/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)
    
    search = request.GET.get('search', '').strip()
    status = request.GET.get('status', '').strip()
    skill = request.GET.get('skill', '').strip()

    qs = Technician.objects.filter(deleted=False).select_related('user').prefetch_related('skills').order_by('-created_at')

    if search:
        qs = qs.filter(
            Q(first_name__icontains=search) |
            Q(last_name__icontains=search) |
            Q(email__icontains=search) |
            Q(phone_number__icontains=search) |
            Q(about__icontains=search) |
            Q(skills__name__icontains=search) |
            Q(user__username__icontains=search) |
            Q(user__first_name__icontains=search) |
            Q(user__last_name__icontains=search)
        ).distinct()

    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)

    if skill:
        qs = qs.filter(Q(skills__name__iexact=skill) | Q(about__icontains=skill)).distinct()

    total_technicians_count = Technician.objects.filter(deleted=False).count()
    filtered_total_count = qs.count()

    try:
        per_page = int(request.GET.get('per_page', 10))
        if per_page <= 0:
            per_page = 10
    except (ValueError, TypeError):
        per_page = 10

    page_num = request.GET.get('page', 1)
    paginator = Paginator(qs, per_page)
    try:
        data_list = paginator.page(page_num)
    except PageNotAnInteger:
        data_list = paginator.page(1)
    except EmptyPage:
        data_list = paginator.page(paginator.num_pages)

    last_page = paginator.num_pages
    current_page = data_list.number
    start_p = max(1, current_page - 2)
    end_p = min(last_page, current_page + 2)
    paginator_list = list(range(start_p, end_p + 1))
    all_pages = list(range(1, last_page + 1))

    common_skills = list(
        Skill.objects.filter(deleted=False, is_active=True)
        .order_by('name')
        .values_list('name', flat=True)
    )

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'all_pages': all_pages,
        'last_page': last_page,
        'per_page': per_page,
        'total_technicians_count': total_technicians_count,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_status': status,
        'selected_skill': skill,
        'common_skills': common_skills,
        'status_choices': ['Active', 'Inactive'],
    }
    return render(request, 'technician/list.html', context) 


@login_required
def technician_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/technician/'):
        messages.error(request, 'You do not have permission to add a technician.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        user_id = request.POST.get('user_id')
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        email = request.POST.get('email', '').strip()
        phone_number = request.POST.get('phone_number', '').strip()
        about = request.POST.get('about', '').strip() 
        
        if not user_id and not first_name:
            messages.error(request, 'Please provide either a User Account or a First Name.')
        else:
            selected_user = None
            if user_id:
                try:
                    selected_user = User.objects.get(id=user_id)
                    if Technician.objects.filter(user=selected_user, deleted=False).exists():
                        messages.error(request, 'This user is already a technician.')
                        return redirect('backend:technician_add')
                except User.DoesNotExist:
                    messages.error(request, 'Selected user does not exist.')
                    return redirect('backend:technician_add')
            
            obj = Technician(
                user=selected_user,
                first_name=first_name,
                last_name=last_name,
                email=email,
                phone_number=phone_number,
                about=about,
                created_by=request.user,
                is_active=True, 
            )
            obj.save()
            skill_ids = [sid for sid in request.POST.getlist('skills') if str(sid).isdigit()]
            if skill_ids:
                obj.skills.set(Skill.objects.filter(id__in=skill_ids, deleted=False))
            messages.success(request, 'Technician added successfully.')
            return redirect('backend:technician_list')
     
    users = User.objects.filter(is_active=True, technician_profile__isnull=True)
    all_skills = Skill.objects.filter(deleted=False, is_active=True).order_by('name')
    context = {
        'users': users,
        'all_skills': all_skills,
    }
    
    return render(request, 'technician/add.html', context)


@login_required
def technician_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/technician/'):
        messages.error(request, 'You do not have permission to edit this technician.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Technician.objects.prefetch_related('skills'), pk=data_id, deleted=False)
    
    if request.method == 'POST':
        user_id = request.POST.get('user_id')
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        email = request.POST.get('email', '').strip()
        phone_number = request.POST.get('phone_number', '').strip()
        about = request.POST.get('about', '').strip() 
        
        if user_id:
            try:
                obj.user = User.objects.get(id=user_id)
            except User.DoesNotExist:
                pass
        
        obj.first_name = first_name
        obj.last_name = last_name
        obj.email = email
        obj.phone_number = phone_number
        obj.about = about
        obj.updated_by = request.user
        obj.save()
        skill_ids = [sid for sid in request.POST.getlist('skills') if str(sid).isdigit()]
        obj.skills.set(Skill.objects.filter(id__in=skill_ids, deleted=False))
        messages.success(request, 'Technician updated successfully.')
        return redirect('backend:technician_list')

    users = User.objects.filter(is_active=True, technician_profile__isnull=True)
    all_skills = Skill.objects.filter(deleted=False, is_active=True).order_by('name')
    selected_skill_ids = set(obj.skills.values_list('id', flat=True))
    context = {
        'obj': obj,
        'users': users,
        'all_skills': all_skills,
        'selected_skill_ids': selected_skill_ids,
    }
    return render(request, 'technician/update.html', context)


@login_required
def technician_details(request, data_id):
    if not checkUserPermission(request, 'can_view', '/backend/technician/'):
        messages.error(request, 'You do not have permission to view this technician.')
        return render(request, '403.html', status=403)

    tech = get_object_or_404(
        Technician.objects.select_related('user').prefetch_related('skills'),
        pk=data_id,
        deleted=False,
    )

    # Approved / completed Service History records for this technician
    rec_filter = Q(assigned_to=tech)
    if tech.user_id:
        rec_filter |= Q(technician_id=tech.user_id)
    all_records = (
        MaintenanceRecord.objects.filter(deleted=False)
        .filter(rec_filter)
        .exclude(ticket__isnull=False, review_status__in=['pending', 'rejected'])
        .select_related('equipment', 'equipment__building', 'equipment__equipment_type', 'ticket', 'maintenance')
        .order_by('-maintenance_date', '-created_at')
    )

    # Assigned Tickets for this technician
    assigned_links = list(
        AssignTechnician.objects.filter(technician=tech, deleted=False, ticket__deleted=False)
        .select_related('ticket', 'ticket__building', 'ticket__equipment', 'ticket__equipment__equipment_type', 'assign_by')
        .order_by('-assigned_at')
    )

    assigned_ticket_ids = [link.ticket_id for link in assigned_links]

    # Map AssignActivities by ticket_id for this technician
    activity_map = {}
    if assigned_ticket_ids:
        for act in AssignActivities.objects.filter(
            technician=tech, ticket_id__in=assigned_ticket_ids, is_active=True
        ).order_by('-updated_at', '-id'):
            if act.ticket_id not in activity_map:
                activity_map[act.ticket_id] = act

    # Map MaintenanceRecords by ticket_id
    ticket_records_map = {}
    if assigned_ticket_ids:
        for mr in MaintenanceRecord.objects.filter(
            ticket_id__in=assigned_ticket_ids, deleted=False
        ).order_by('-created_at', '-id'):
            ticket_records_map.setdefault(mr.ticket_id, []).append(mr)

    # Map checklist counts by ticket_id
    checklist_counts_map = {}
    if assigned_ticket_ids:
        for t_item in TicketTaskItem.objects.filter(
            ticket_id__in=assigned_ticket_ids, deleted=False, is_active=True
        ).only('ticket_id', 'is_completed'):
            entry = checklist_counts_map.setdefault(t_item.ticket_id, [0, 0])
            entry[0] += 1
            if t_item.is_completed:
                entry[1] += 1

    # Build enriched assigned task rows with acceptance + stage status
    all_task_rows = []
    pending_acceptance_count = 0
    accepted_ready_count = 0
    in_progress_tasks_count = 0
    active_work_count = 0
    pending_review_count = 0
    resolved_tickets = 0

    for link in assigned_links:
        tkt = link.ticket
        act = activity_map.get(tkt.id)
        tkt_records = ticket_records_map.get(tkt.id, [])
        latest_rec = tkt_records[0] if tkt_records else None
        approved_rec = next((r for r in tkt_records if r.review_status == 'approved'), None)
        pending_rec = next((r for r in tkt_records if r.review_status == 'pending'), None)

        raw_accept = (act.status if act else 'pending') or 'pending'
        if raw_accept == 'completed':
            acceptance_status = 'accepted'
        elif raw_accept in ('pending', 'accepted', 'rejected'):
            acceptance_status = raw_accept
        else:
            acceptance_status = 'pending'

        completion_status = (act.completion_status if act else 'pending') or 'pending'

        if acceptance_status == 'rejected':
            stage_key = 'rejected'
            stage_label = 'Rejected'
        elif tkt.status in ('resolved', 'closed') or approved_rec:
            stage_key = 'completed'
            stage_label = 'Completed & Approved'
            resolved_tickets += 1
        elif tkt.status == 'pending_review' or pending_rec:
            stage_key = 'pending_review'
            stage_label = 'Pending Review'
            pending_review_count += 1
            active_work_count += 1
        elif tkt.status == 'in_progress' or completion_status == 'in_progress':
            stage_key = 'in_progress'
            stage_label = 'In Progress'
            in_progress_tasks_count += 1
            active_work_count += 1
        elif acceptance_status == 'accepted' or tkt.status == 'accepted':
            stage_key = 'accepted'
            stage_label = 'Accepted'
            accepted_ready_count += 1
            active_work_count += 1
        else:
            stage_key = 'assigned'
            stage_label = 'Pending Acceptance'
            pending_acceptance_count += 1

        is_terminal = stage_key in ('completed', 'rejected') or tkt.status in ('resolved', 'closed', 'rejected')
        can_accept = (acceptance_status == 'pending') and not is_terminal
        can_reject = can_accept
        can_start = (acceptance_status == 'accepted') and (stage_key == 'accepted')
        can_submit = (acceptance_status == 'accepted') and (stage_key == 'in_progress') and bool(tkt.equipment_id)
        needs_rework = bool(latest_rec and latest_rec.review_status == 'rejected' and stage_key == 'in_progress')

        submission_count = len(tkt_records)
        revision_count = sum(1 for r in tkt_records if r.review_status == 'rejected')
        review_count = sum(1 for r in tkt_records if r.review_status in ('approved', 'rejected'))
        latest_rejected = next((r for r in tkt_records if r.review_status == 'rejected'), None)
        revision_notes = (latest_rejected.review_notes or '').strip() if latest_rejected else ''
        prefill_work_desc = (
            latest_rejected.work_description
            if (needs_rework and latest_rejected and latest_rejected.work_description)
            else (tkt.issue_description or '')
        )

        cl_total, cl_done = checklist_counts_map.get(tkt.id, (0, 0))

        all_task_rows.append({
            'assignment': link,
            'ticket': tkt,
            'activity': act,
            'acceptance_status': acceptance_status,
            'completion_status': completion_status,
            'stage_key': stage_key,
            'stage_label': stage_label,
            'can_accept': can_accept,
            'can_reject': can_reject,
            'can_start': can_start,
            'can_submit': can_submit,
            'needs_rework': needs_rework,
            'checklist_total': cl_total,
            'checklist_done': cl_done,
            'submission_count': submission_count,
            'revision_count': revision_count,
            'review_count': review_count,
            'latest_rejected': latest_rejected,
            'revision_notes': revision_notes,
            'prefill_work_desc': prefill_work_desc,
            'latest_record': latest_rec,
            'approved_record': approved_rec,
            'pending_record': pending_rec,
        })

    # Service Overview combines active/completed assigned ticket tasks + non-ticket (scheduled/manual) records without double-counting
    non_ticket_records = all_records.filter(ticket__isnull=True)
    completed_services = resolved_tickets + non_ticket_records.filter(maintainance_status='completed').count()
    in_progress_services = in_progress_tasks_count + non_ticket_records.filter(maintainance_status='in_progress').count()
    pending_review_services = pending_review_count
    pending_services = pending_acceptance_count + accepted_ready_count + non_ticket_records.filter(maintainance_status='pending').count()
    total_services = completed_services + in_progress_services + pending_review_services + pending_services
    approved_services = all_records.filter(review_status='approved').count()
    history_total_count = all_records.count()

    total_tickets = len(all_task_rows)
    open_tickets = pending_acceptance_count + active_work_count

    # Active tab: 'tasks' (default - Assigned Tasks) or 'history' (Service History)
    active_tab = (request.GET.get('tab') or 'tasks').strip()
    if active_tab not in ('tasks', 'history'):
        active_tab = 'tasks'

    try:
        per_page = int(request.GET.get('per_page', 10))
        if per_page not in (10, 20, 50, 100):
            per_page = 10
    except (TypeError, ValueError):
        per_page = 10

    # Filters for Assigned Tasks tab
    selected_acceptance = (request.GET.get('acceptance') or '').strip()
    selected_stage = (request.GET.get('stage') or '').strip()
    filtered_tasks = all_task_rows
    if selected_acceptance:
        filtered_tasks = [r for r in filtered_tasks if r['acceptance_status'] == selected_acceptance]
    if selected_stage:
        filtered_tasks = [r for r in filtered_tasks if r['stage_key'] == selected_stage]

    # Filters for Service History tab
    selected_type = (request.GET.get('type') or '').strip()
    selected_status = (request.GET.get('status') or '').strip()
    filtered_records = all_records
    if selected_type:
        filtered_records = filtered_records.filter(record_type=selected_type)
    if selected_status:
        filtered_records = filtered_records.filter(maintainance_status=selected_status)

    page_num = request.GET.get('page', 1)
    if active_tab == 'history':
        filtered_count = filtered_records.count()
        paginator = Paginator(filtered_records, per_page)
        page_obj = paginator.get_page(page_num)
        task_page_obj = Paginator(filtered_tasks, per_page).get_page(1)
        service_records_page = page_obj
    else:
        filtered_count = len(filtered_tasks)
        paginator = Paginator(filtered_tasks, per_page)
        page_obj = paginator.get_page(page_num)
        task_page_obj = page_obj
        service_records_page = Paginator(filtered_records, per_page).get_page(1)

    task_stage_choices = [
        ('assigned', 'Pending Acceptance'),
        ('accepted', 'Accepted'),
        ('in_progress', 'In Progress'),
        ('pending_review', 'Pending Review'),
        ('completed', 'Completed & Approved'),
        ('rejected', 'Rejected'),
    ]
    acceptance_choices = [
        ('pending', 'Pending Acceptance'),
        ('accepted', 'Accepted'),
        ('rejected', 'Rejected'),
    ]

    context = {
        'tech': tech,
        'skills_list': tech.active_skills,
        'active_tab': active_tab,
        'assigned_tasks': task_page_obj,
        'tasks_total_count': len(all_task_rows),
        'filtered_tasks_count': len(filtered_tasks),
        'history_total_count': history_total_count,
        'service_records': service_records_page,
        'page_obj': page_obj,
        'per_page': per_page,
        'filtered_count': filtered_count,
        'assigned_links': assigned_links[:10],
        'total_services': total_services,
        'completed_services': completed_services,
        'in_progress_services': in_progress_services,
        'pending_review_services': pending_review_services,
        'pending_services': pending_services,
        'approved_services': approved_services,
        'total_tickets': total_tickets,
        'pending_acceptance_count': pending_acceptance_count,
        'active_work_count': active_work_count,
        'pending_review_count': pending_review_count,
        'open_tickets': open_tickets,
        'resolved_tickets': resolved_tickets,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'maint_status_choices': MaintenanceRecord.MAINTAIANCE_STATUS_CHOICES,
        'task_stage_choices': task_stage_choices,
        'acceptance_choices': acceptance_choices,
        'selected_acceptance': selected_acceptance,
        'selected_stage': selected_stage,
        'selected_type': selected_type,
        'selected_status': selected_status,
        'today_iso': timezone.localdate().strftime('%Y-%m-%d'),
    }
    return render(request, 'technician/details.html', context)


@login_required
def technician_task_detail(request, tech_id, ticket_id):
    """Dedicated Technician Task Details page reached after accepting + QR & PIN verification."""
    if not checkUserPermission(request, 'can_view', '/backend/technician/'):
        messages.error(request, 'You do not have permission to view this task.')
        return render(request, '403.html', status=403)

    tech = get_object_or_404(Technician, pk=tech_id, deleted=False)
    ticket = get_object_or_404(
        Ticket.objects.select_related('equipment', 'equipment__equipment_type', 'building', 'user'),
        pk=ticket_id,
        deleted=False,
    )

    assignment = AssignTechnician.objects.filter(
        ticket=ticket, technician=tech, deleted=False
    ).first()
    if not assignment:
        messages.error(request, 'This ticket is not assigned to this technician.')
        return redirect('backend:technician_details', data_id=tech.id)

    act = AssignActivities.objects.filter(
        ticket=ticket, technician=tech, is_active=True
    ).order_by('-updated_at', '-id').first()
    raw_acc = (act.status if act else 'pending') or 'pending'
    acceptance_status = 'accepted' if raw_acc == 'completed' else raw_acc

    # Gate 1: Must accept task on Technician Profile first
    if acceptance_status != 'accepted':
        messages.warning(request, 'Please accept this assigned task first before opening task details.')
        return redirect(f"{reverse('backend:technician_details', kwargs={'data_id': tech.id})}?tab=tasks")

    # Gate 2: Must pass through Equipment QR + PIN page before first entry when still in 'accepted' stage
    pin_session_key = f'ticket_pin_verified_{ticket.id}_{tech.id}'
    is_pin_verified = bool(
        request.session.get(pin_session_key)
        or (act and act.started_at)
        or ticket.status in ('in_progress', 'pending_review', 'resolved', 'closed')
    )
    if request.method == 'GET' and not is_pin_verified and ticket.equipment_id:
        qr_url = reverse('backend:equipment_qr', kwargs={'data_id': ticket.equipment_id})
        return redirect(f"{qr_url}?view=1&ticket={ticket.id}&technician={tech.id}")

    checklist_items = list(
        ticket.checklist_items.filter(deleted=False, is_active=True).order_by('order', 'id')
    )
    service_records = list(
        MaintenanceRecord.objects.filter(ticket=ticket, deleted=False)
        .select_related('assigned_to', 'technician', 'reviewed_by')
        .prefetch_related('components')
        .order_by('-created_at', '-id')
    )
    pending_record = next((r for r in service_records if r.review_status == 'pending'), None)
    approved_record = next((r for r in service_records if r.review_status == 'approved'), None)
    latest_rejected = next((r for r in service_records if r.review_status == 'rejected'), None)

    submission_count = len(service_records)
    revision_count = sum(1 for r in service_records if r.review_status == 'rejected')
    review_count = sum(1 for r in service_records if r.review_status in ('approved', 'rejected'))

    is_editable = (ticket.status == 'in_progress') and not approved_record and not pending_record
    needs_rework = bool(latest_rejected and is_editable and service_records and service_records[0].review_status == 'rejected')

    if request.method == 'POST':
        if not is_editable:
            messages.warning(request, 'This task is no longer open for editing.')
            return redirect('backend:technician_task_detail', tech_id=tech.id, ticket_id=ticket.id)

        action = (request.POST.get('action') or 'submit_work').strip()
        completed_ids = set(request.POST.getlist('completed_task_ids'))
        now = timezone.now()

        # Persist checklist ticks + per-item feedback
        done_count = 0
        for item in checklist_items:
            checked = str(item.id) in completed_ids
            fb = (request.POST.get(f'item_feedback_{item.id}') or '').strip()
            item.is_completed = checked
            item.technician_feedback = fb
            item.completed_at = now if checked else None
            item.updated_by = request.user
            item.save()
            if checked:
                done_count += 1

        if action == 'save_progress':
            messages.success(request, f'Checklist progress saved ({done_count}/{len(checklist_items)} completed).')
            return redirect('backend:technician_task_detail', tech_id=tech.id, ticket_id=ticket.id)

        if action == 'submit_work':
            if not ticket.equipment:
                messages.error(request, 'Cannot submit work: this ticket has no linked equipment.')
                return redirect('backend:technician_task_detail', tech_id=tech.id, ticket_id=ticket.id)

            work_description = (request.POST.get('work_description') or '').strip()
            if not work_description:
                if checklist_items:
                    work_description = f"Completed {done_count} of {len(checklist_items)} assigned checklist tasks."
                else:
                    work_description = ticket.issue_description or ticket.title

            record_type = (request.POST.get('record_type') or 'repair').strip() or 'repair'
            maint_date_str = (request.POST.get('maintenance_date') or '').strip()
            parts_replaced = (request.POST.get('parts_replaced') or '').strip()
            cost_raw = (request.POST.get('cost') or '').strip()
            duration_raw = (request.POST.get('duration_hours') or '').strip()

            try:
                maint_date = datetime.strptime(maint_date_str, '%Y-%m-%d').date() if maint_date_str else timezone.localdate()
            except ValueError:
                maint_date = timezone.localdate()

            cost_val = None
            if cost_raw:
                try:
                    cost_val = float(cost_raw)
                except ValueError:
                    cost_val = None

            duration_val = None
            if duration_raw:
                try:
                    duration_val = float(duration_raw)
                except ValueError:
                    duration_val = None

            record = MaintenanceRecord.objects.create(
                equipment=ticket.equipment,
                ticket=ticket,
                record_type=record_type,
                maintenance_date=maint_date,
                assigned_to=tech,
                technician=tech.user or request.user,
                work_description=work_description,
                parts_replaced=parts_replaced or '',
                cost=cost_val,
                duration_hours=duration_val,
                maintainance_status='completed',
                review_status='pending',
                created_by=request.user,
            )

            # Snapshot checklist items + feedback onto the submitted MaintenanceRecord
            for item in checklist_items:
                MaintenanceComponent.objects.create(
                    maintenance_record=record,
                    name=item.title[:100],
                    description=item.title,
                    is_checked=item.is_completed,
                    remark=item.technician_feedback or '',
                    proceed=item.is_completed,
                    created_by=request.user,
                )

            ticket.status = 'pending_review'
            ticket.updated_by = request.user
            ticket.save()

            AssignActivities.objects.filter(ticket=ticket, is_active=True, status='accepted').update(
                completion_status='submitted',
                updated_by=request.user,
            )

            submission_num = MaintenanceRecord.objects.filter(ticket=ticket, deleted=False).count()
            prior_revisions = MaintenanceRecord.objects.filter(
                ticket=ticket, deleted=False, review_status='rejected'
            ).count()
            checklist_summary = f" [{done_count}/{len(checklist_items)} tasks checked]" if checklist_items else ""
            if submission_num > 1 and prior_revisions > 0:
                sub_prefix = f"Submission #{submission_num} (Updated after Revision #{prior_revisions}) — "
                msg_text = f'Updated work (Submission #{submission_num}) submitted for supervisor review.'
            else:
                sub_prefix = f"Submission #{submission_num} — "
                msg_text = 'Work submitted for supervisor review. Once approved, it will move to Service History.'

            TicketActivityLog.objects.create(
                ticket=ticket,
                actor=request.user,
                action_type='submitted_for_review',
                description=(
                    f"{sub_prefix}{tech.display_name} submitted work{checklist_summary} "
                    f"for supervisor review (SR-{record.id:03d}): {work_description}"
                ),
                metadata={
                    'record_id': record.id,
                    'technician_id': tech.id,
                    'submission_num': submission_num,
                    'prior_revisions': prior_revisions,
                    'checklist_done': done_count,
                    'checklist_total': len(checklist_items),
                },
            )
            messages.success(request, msg_text)
            return redirect(f"{reverse('backend:technician_details', kwargs={'data_id': tech.id})}?tab=tasks")

    checklist_total = len(checklist_items)
    checklist_done = sum(1 for i in checklist_items if i.is_completed)
    checklist_pct = int(round((checklist_done / checklist_total) * 100)) if checklist_total else 0

    prefill_desc = ''
    if needs_rework and latest_rejected and latest_rejected.work_description:
        prefill_desc = latest_rejected.work_description
    elif pending_record:
        prefill_desc = pending_record.work_description
    elif approved_record:
        prefill_desc = approved_record.work_description

    chronological_records = list(reversed(service_records))
    numbered_rejected_records = []
    rev_counter = 0
    for idx, rec in enumerate(chronological_records, start=1):
        if rec.review_status == 'rejected':
            rev_counter += 1
            numbered_rejected_records.append({
                'record': rec,
                'attempt_num': idx,
                'revision_num': rev_counter,
            })
    numbered_rejected_records.reverse()

    context = {
        'tech': tech,
        'ticket': ticket,
        'assignment': assignment,
        'activity': act,
        'checklist_items': checklist_items,
        'checklist_total': checklist_total,
        'checklist_done': checklist_done,
        'checklist_pct': checklist_pct,
        'is_editable': is_editable,
        'needs_rework': needs_rework,
        'latest_rejected': latest_rejected,
        'pending_record': pending_record,
        'approved_record': approved_record,
        'submission_count': submission_count,
        'next_submission_num': submission_count + 1,
        'revision_count': revision_count,
        'review_count': review_count,
        'numbered_rejected_records': numbered_rejected_records,
        'prefill_desc': prefill_desc,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'today_iso': timezone.localdate().strftime('%Y-%m-%d'),
    }
    return render(request, 'technician/task_detail.html', context)


@login_required
def skill_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/technician/'):
        messages.error(request, 'You do not have permission to view skill management.')
        return render(request, '403.html', status=403)

    search = (request.GET.get('search') or '').strip()
    status = (request.GET.get('status') or '').strip()

    qs = (
        Skill.objects.filter(deleted=False)
        .annotate(technicians_count=Count('technicians', filter=Q(technicians__deleted=False), distinct=True))
        .order_by('name')
    )

    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(description__icontains=search))
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)

    total_skills_count = Skill.objects.filter(deleted=False).count()
    filtered_total_count = qs.count()

    try:
        per_page = int(request.GET.get('per_page', 10))
        if per_page not in (10, 20, 50, 100):
            per_page = 10
    except (TypeError, ValueError):
        per_page = 10

    paginator = Paginator(qs, per_page)
    data_list = paginator.get_page(request.GET.get('page', 1))

    context = {
        'data_list': data_list,
        'per_page': per_page,
        'total_skills_count': total_skills_count,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_status': status,
        'status_choices': ['Active', 'Inactive'],
    }
    return render(request, 'technician/skills.html', context)


@login_required
def skill_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/technician/'):
        messages.error(request, 'You do not have permission to add skills.')
        return render(request, '403.html', status=403)

    if request.method == 'POST':
        name = (request.POST.get('name') or '').strip()
        description = (request.POST.get('description') or '').strip()
        status_val = (request.POST.get('status') or 'Active').strip()
        is_active = (status_val != 'Inactive')

        if not name:
            messages.error(request, 'Skill name is required.')
        elif Skill.objects.filter(name__iexact=name, deleted=False).exists():
            messages.error(request, f'Skill "{name}" already exists.')
        else:
            existing_deleted = Skill.objects.filter(name__iexact=name, deleted=True).first()
            if existing_deleted:
                existing_deleted.name = name
                existing_deleted.description = description
                existing_deleted.is_active = is_active
                existing_deleted.deleted = False
                existing_deleted.updated_by = request.user
                existing_deleted.save()
            else:
                Skill.objects.create(
                    name=name,
                    description=description,
                    is_active=is_active,
                    created_by=request.user,
                    updated_by=request.user,
                )
            messages.success(request, f'Skill "{name}" added successfully.')
    return redirect('backend:skill_list')


@login_required
def skill_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/technician/'):
        messages.error(request, 'You do not have permission to edit skills.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Skill, pk=data_id, deleted=False)
    if request.method == 'POST':
        name = (request.POST.get('name') or '').strip()
        description = (request.POST.get('description') or '').strip()
        status_val = (request.POST.get('status') or ('Active' if obj.is_active else 'Inactive')).strip()

        if not name:
            messages.error(request, 'Skill name is required.')
        elif Skill.objects.filter(name__iexact=name, deleted=False).exclude(pk=obj.pk).exists():
            messages.error(request, f'Skill "{name}" already exists.')
        else:
            obj.name = name
            obj.description = description
            obj.is_active = (status_val != 'Inactive')
            obj.updated_by = request.user
            obj.save()
            messages.success(request, f'Skill "{name}" updated successfully.')
    return redirect('backend:skill_list')


@login_required
def skill_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/technician/'):
        messages.error(request, 'You do not have permission to change skill status.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Skill, pk=data_id, deleted=False)
    if request.method == 'POST':
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        messages.success(request, f'Skill "{obj.name}" {"activated" if obj.is_active else "deactivated"}.')
    return redirect('backend:skill_list')


@login_required
def skill_delete(request, data_id):
    if not checkUserPermission(request, 'can_delete', '/backend/technician/'):
        messages.error(request, 'You do not have permission to delete skills.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Skill, pk=data_id, deleted=False)
    obj.deleted = True
    obj.is_active = False
    obj.updated_by = request.user
    obj.save()
    messages.success(request, f'Skill "{obj.name}" deleted successfully.')
    return redirect('backend:skill_list')


@login_required
def technician_delete(request, data_id):
    if not checkUserPermission(request, 'can_delete', '/backend/technician/'):
        messages.error(request, 'You do not have permission to delete this technician.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Technician, pk=data_id, deleted=False)
    obj.deleted = True
    obj.deleted_by = request.user
    obj.updated_at = timezone.now()
    obj.save()
    messages.success(request, 'Technician deleted successfully.')
    return redirect('backend:technician_list')





# Equipment Management 
@login_required 
def equipment_type_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/equipment/equipment_type/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)
    
    search = request.GET.get('search', '').strip()
    status = request.GET.get('status', '').strip()
    
    qs = EquipmentType.objects.filter(deleted=False).annotate(
        equipment_count=Count('equipment', filter=Q(equipment__deleted=False))
    ).order_by('id')
    
    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(description__icontains=search))
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count
    
    context = {
        'data_list': data_list,
        'equipment_types': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_status': status,
        'status_choices': ['Active', 'Inactive'],
    }
    return render(request, 'equipment_type/list.html', context)


@login_required
def equipment_type_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/equipment/equipment_type/'):
        messages.error(request, 'You do not have permission to add this equipment type.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        description = request.POST.get('description', '').strip()
        is_active = request.POST.get('is_active', '1') != '0'
        
        if not name:
            messages.error(request, 'Name is required.')
        else:
            obj = EquipmentType(
                name=name,
                description=description,
                is_active=is_active,
                created_by=request.user
            )
            obj.save()
            messages.success(request, 'Equipment Type added successfully.')
            return redirect('backend:equipment_type_list')
    
    return render(request, 'equipment_type/add.html', {})


@login_required
def equipment_type_update(request, data_id):
    if not checkUserPermission(request, 'can_edit', '/backend/equipment/equipment_type/'):
        messages.error(request, 'You do not have permission to edit this equipment type.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(EquipmentType, pk=data_id, deleted=False)
    
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        description = request.POST.get('description', '').strip() 
        
        if not name:
            messages.error(request, 'Name is required.')
        else:
            obj.name = name
            obj.description = description
            if 'is_active' in request.POST:
                obj.is_active = request.POST.get('is_active') != '0'
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Equipment Type updated successfully.')
            return redirect('backend:equipment_type_list')
    
    return render(request, 'equipment_type/update.html', {'obj': obj})


@login_required
def equipment_type_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/equipment/equipment_type/'):
        messages.error(request, 'You do not have permission to change status of this equipment type.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        obj = get_object_or_404(EquipmentType, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Equipment Type status changed to {status_text} successfully.')
        return redirect('backend:equipment_type_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:equipment_type_list')


@login_required
def equipment_type_data_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/equipment/equipment_type_data/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)
    
    search = request.GET.get('search', '').strip()
    type_id = request.GET.get('type_id', '').strip()
    status = request.GET.get('status', '').strip()
    
    qs = EquipmentTypeData.objects.filter(deleted=False).select_related('equipment_type').order_by('equipment_type__name', 'id')
    
    if search:
        qs = qs.filter(Q(field_name__icontains=search) | Q(equipment_type__name__icontains=search))
    if type_id:
        qs = qs.filter(equipment_type_id=type_id)
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')
    
    context = {
        'data_list': data_list,
        'equipment_type_datas': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'type_id': type_id,
        'selected_status': status,
        'equipment_types': equipment_types,
        'status_choices': ['Active', 'Inactive'],
    }
    return render(request, 'equipment_type_data/list.html', context)


@login_required
def equipment_type_data_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/equipment/equipment_type_data/'):
        messages.error(request, 'You do not have permission to add this equipment type data.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        equipment_type_id = request.POST.get('equipment_type_id', '').strip()
        field_name = request.POST.get('field_name', '').strip()
        field_type = request.POST.get('field_type', '').strip()
        is_required = request.POST.get('is_required', '') in ('True', 'true', '1', 'on')
        is_active = request.POST.get('is_active', '1') != '0'
        
        if not equipment_type_id or not field_name or not field_type:
            messages.error(request, 'Equipment Type, Field Name, and Field Type are required.')
        else:
            obj = EquipmentTypeData(
                equipment_type_id=equipment_type_id,
                field_name=field_name,
                field_type=field_type,
                is_required=is_required,
                is_active=is_active,
                created_by=request.user
            )
            obj.save()
            messages.success(request, 'Equipment Type Data added successfully.')
            return redirect('backend:equipment_type_data_list')
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    return render(request, 'equipment_type_data/add.html', {'equipment_types': equipment_types})


@login_required
def equipment_type_data_update(request, data_id):
    if not checkUserPermission(request, 'can_edit', '/backend/equipment/equipment_type_data/'):
        messages.error(request, 'You do not have permission to edit this equipment type data.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(EquipmentTypeData, pk=data_id, deleted=False)
    
    if request.method == 'POST':
        equipment_type_id = request.POST.get('equipment_type_id', '').strip()
        field_name = request.POST.get('field_name', '').strip()
        field_type = request.POST.get('field_type', '').strip()
        is_required = request.POST.get('is_required', '') in ('True', 'true', '1', 'on')
      
        if not equipment_type_id or not field_name or not field_type:
            messages.error(request, 'Equipment Type, Field Name, and Field Type are required.')
        else:
            obj.equipment_type_id = equipment_type_id
            obj.field_name = field_name
            obj.field_type = field_type
            obj.is_required = is_required
            if 'is_active' in request.POST:
                obj.is_active = request.POST.get('is_active') != '0'
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Equipment Type Data updated successfully.')
            return redirect('backend:equipment_type_data_list')
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    return render(request, 'equipment_type_data/update.html', {'obj': obj, 'equipment_types': equipment_types})


@login_required
def equipment_type_data_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/equipment/equipment_type_data/'):
        messages.error(request, 'You do not have permission to change status of this equipment type data.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        obj = get_object_or_404(EquipmentTypeData, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Equipment Type Data status changed to {status_text} successfully.')
        return redirect('backend:equipment_type_data_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:equipment_type_data_list')



@login_required
def equipment_component_preset_list(request):
    if not checkUserPermission(request, "can_view", '/backend/component-preset/'):
        messages.error(request, "You do not have permission to view this page.")
        return render(request, '403.html', status=403)
    
    search = request.GET.get('search', '').strip()
    equipment_type_id = request.GET.get('equipment_type', '').strip()
    status = request.GET.get('status', '').strip()
    
    qs = EquipmentComponentPreset.objects.filter(deleted=False).select_related('equipment_type').order_by('id')
    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(description__icontains=search) | Q(equipment_type__name__icontains=search))
    if equipment_type_id:
        qs = qs.filter(equipment_type_id=equipment_type_id)
    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')
    
    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'equipment_type_id': equipment_type_id,
        'selected_status': status,
        'equipment_types': equipment_types,
        'status_choices': ['Active', 'Inactive'],
    }
    return render(request, 'equipment-preset/list.html', context)


@login_required
def equipment_component_preset_add(request):
    if not checkUserPermission(request, "can_add", '/backend/component-preset/'):
        messages.error(request, "You do not have permission to add this page.")
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        equipment_type_id = request.POST.get('equipment_type', '').strip()
        description = request.POST.get('description', '').strip()
        is_active = request.POST.get('is_active', '1') != '0'
        
        if not name or not equipment_type_id:
            messages.error(request, 'Name and Equipment Type are required.')
        else:
            EquipmentComponentPreset.objects.create(
                name=name,
                equipment_type_id=equipment_type_id,
                description=description,
                created_by=request.user,
                is_active=is_active
            )
            messages.success(request, 'Component Preset added successfully.')
            return redirect('backend:equipment_component_preset_list')
            
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    context = {
        'action': 'Add',
        'equipment_types': equipment_types,
    }
    return render(request, 'equipment-preset/add.html', context)


@login_required
def equipment_component_preset_update(request, data_id):
    if not checkUserPermission(request, "can_update", '/backend/component-preset/'):
        messages.error(request, "You do not have permission to update this page.")
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(EquipmentComponentPreset, pk=data_id, deleted=False)
    
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        equipment_type_id = request.POST.get('equipment_type', '').strip()
        description = request.POST.get('description', '').strip()
        
        if not name or not equipment_type_id:
            messages.error(request, 'Name and Equipment Type are required.')
        else:
            obj.name = name
            obj.equipment_type_id = equipment_type_id
            obj.description = description
            if 'is_active' in request.POST:
                obj.is_active = request.POST.get('is_active') != '0'
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Component Preset updated successfully.')
            return redirect('backend:equipment_component_preset_list')
            
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    context = {
        'action': 'Update',
        'obj': obj,
        'equipment_types': equipment_types,
    }
    return render(request, 'equipment-preset/update.html', context)


@login_required
def equipment_component_preset_delete(request, data_id):
    if not checkUserPermission(request, "can_delete", '/backend/component-preset/'):
        messages.error(request, "You do not have permission to delete this page.")
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(EquipmentComponentPreset, pk=data_id, deleted=False)
    obj.deleted = True
    obj.updated_by = request.user
    obj.save()
    messages.success(request, 'Component Preset deleted successfully.')
    return redirect('backend:equipment_component_preset_list')


@login_required
def equipment_component_preset_status(request, data_id):
    if not checkUserPermission(request, "can_update", '/backend/component-preset/'):
        messages.error(request, "You do not have permission to update this page.")
        return render(request, '403.html', status=403)
        
    if request.method == 'POST':
        obj = get_object_or_404(EquipmentComponentPreset, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Component Preset status changed to {status_text} successfully.')
    return redirect('backend:equipment_component_preset_list')
    
   

@login_required
def equipment_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)
    
    search = request.GET.get('search', '').strip()
    type_id = request.GET.get('type_id', '').strip()
    status = request.GET.get('status', '').strip()
    
    qs = Equipment.objects.filter(deleted=False).select_related('equipment_type', 'building').annotate(components_count=Count('components')).order_by('-created_at')
    
    if search:
        qs = qs.filter(Q(equipment_id__icontains=search) | Q(brand__icontains=search) | Q(building__name__icontains=search))
    if type_id:
        qs = qs.filter(equipment_type_id=type_id)
    if status:
        qs = qs.filter(status=status)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('-created_at')
    
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'type_id': type_id,
        'status': status,
        'equipment_types': equipment_types,
        'status_choices': Equipment.STATUS_CHOICES,
    }   
    return render(request, 'equipment/list.html', context)

from backend.models import EquipmentComponents 

@login_required
@require_POST
def quick_create_equipment_type(request):
    """AJAX endpoint to quick-create an equipment type from the searchable dropdown"""
    name = request.POST.get('name', '').strip()
    if not name:
        return JsonResponse({'success': False, 'error': 'Equipment type name is required.'}, status=400)
    
    existing = EquipmentType.objects.filter(name__iexact=name).first()
    if existing:
        if existing.deleted:
            existing.deleted = False
            existing.is_active = True
            existing.updated_by = request.user
            existing.save()
        return JsonResponse({
            'success': True,
            'id': existing.id,
            'name': existing.name,
            'already_existed': True,
            'fields': []
        })
    
    new_type = EquipmentType.objects.create(
        name=name,
        created_by=request.user,
        is_active=True
    )
    return JsonResponse({
        'success': True,
        'id': new_type.id,
        'name': new_type.name,
        'already_existed': False,
        'fields': []
    })


@login_required
def equipment_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to add equipment.')
        return render(request, '403.html', status=403)
        
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')
    buildings = Building.objects.filter(deleted=False, is_active=True).select_related('division', 'subdivision', 'section').order_by('name')
    
    # Query distinct brands or provide curated defaults
    existing_brands = list(Equipment.objects.filter(deleted=False).exclude(brand='').values_list('brand', flat=True).distinct().order_by('brand'))
    default_brands = ['Carrier', 'Caterpillar', 'Cummins', 'Daikin', 'Honeywell', 'Otis', 'Schindler', 'Siemens']
    all_brands = sorted(list(set(existing_brands + default_brands)))

    # Serialize equipment types with their dynamic data_fields
    equipment_types_data = []
    for et in equipment_types:
        fields = [
            {
                'id': df.id,
                'field_name': df.field_name,
                'field_type': df.field_type,
                'is_required': df.is_required,
            }
            for df in et.data_fields.filter(is_active=True, deleted=False).order_by('id')
        ]
        equipment_types_data.append({
            'id': et.id,
            'name': et.name,
            'fields': fields,
        })
    
    # Serialize buildings with metadata for review step
    buildings_data = []
    for b in buildings:
        sec = b.section.name if b.section else ''
        sub = b.subdivision.name if b.subdivision else ''
        div = b.division.name if b.division else ''
        org_parts = [p for p in [sec, sub, div] if p]
        org_subdivision = " / ".join(org_parts) if org_parts else ""
        buildings_data.append({
            'id': b.id,
            'name': b.name,
            'code': b.building_id or '',
            'org_subdivision': org_subdivision,
        })

    from_building = request.POST.get('from_building') or request.GET.get('building') or request.GET.get('building_id')
    from_building_id = None
    if from_building:
        try:
            from_building_id = int(from_building)
        except (ValueError, TypeError):
            from_building_id = None

    if request.method == 'POST':
        is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.POST.get('is_ajax') == '1'

        equipment_id = request.POST.get('equipment_id', '').strip()
        equipment_name = request.POST.get('equipment_name', '').strip()
        equipment_type_id = request.POST.get('equipment_type_id') or None
        serial_number = request.POST.get('serial_number', '').strip()
        model_number = request.POST.get('model_number', '').strip()
        brand = request.POST.get('brand', '').strip()
        
        # Custom dynamic fields
        custom_fields_raw = request.POST.get('custom_fields', '{}')
        try:
            custom_fields = json.loads(custom_fields_raw) if isinstance(custom_fields_raw, str) else custom_fields_raw
        except Exception:
            custom_fields = {}

        # Building assignment
        assign_to_building = request.POST.get('assign_to_building') in ['true', 'on', '1', True]
        building_id = request.POST.get('building_id') if assign_to_building else None
        floor_location = request.POST.get('floor_location', '').strip() if assign_to_building else ''
        room_location = request.POST.get('room_location', '').strip() if assign_to_building else ''
        installment_date = request.POST.get('installment_date', '').strip() if assign_to_building else None
        purchase_date = request.POST.get('purchase_date', '').strip() or None
        warranty_expire_date = request.POST.get('warranty_expire_date', '').strip() or None
        installation_notes = request.POST.get('installation_notes', '').strip() if assign_to_building else ''

        # Maintenance
        maintenance_period_days = request.POST.get('maintenance_period_days', '90').strip()
        frequency_unit = request.POST.get('frequency_unit', 'Days').strip()
        last_maintenance_date = request.POST.get('last_maintenance_date', '').strip()
        next_service_due = request.POST.get('next_service_due', '').strip()
        maintenance_notes = request.POST.get('maintenance_notes', '').strip()
        status = request.POST.get('status', 'active').strip()

        # Validation
        errors = []
        if not equipment_id:
            errors.append('Equipment ID is required.')
        elif Equipment.objects.filter(equipment_id=equipment_id, deleted=False).exists():
            errors.append(f"Equipment ID '{equipment_id}' already exists. Please use a unique ID.")

        if not equipment_type_id:
            errors.append('Equipment Type is required.')
        if not equipment_name:
            errors.append('Equipment Name is required.')
        if not serial_number:
            errors.append('Serial Number is required.')
        if not model_number:
            errors.append('Model Number is required.')
        if not brand:
            errors.append('Brand is required.')

        if assign_to_building:
            if not building_id:
                errors.append('Building selection is required when assigning to a building.')
            if not installment_date:
                errors.append('Installation Date is required when assigning to a building.')

        if errors:
            if is_ajax:
                return JsonResponse({'success': False, 'errors': errors}, status=400)
            for err in errors:
                messages.error(request, err)
            context = {
                'equipment_types': equipment_types,
                'equipment_types_json': json.dumps(equipment_types_data),
                'buildings': buildings,
                'buildings_json': json.dumps(buildings_data),
                'brands': all_brands,
                'from_building': from_building_id,
            }
            return render(request, 'equipment/add.html', context)

        # Parse numeric and foreign key values safely
        try:
            m_days = int(maintenance_period_days) if maintenance_period_days else 90
        except ValueError:
            m_days = 90

        try:
            b_id = int(building_id) if building_id else None
        except (ValueError, TypeError):
            b_id = None

        try:
            t_id = int(equipment_type_id) if equipment_type_id else None
        except (ValueError, TypeError):
            t_id = None

        obj = Equipment(
            equipment_id=equipment_id,
            equipment_name=equipment_name,
            equipment_type_id=t_id,
            building_id=b_id,
            brand=brand,
            model_number=model_number,
            serial_number=serial_number,
            floor_location=floor_location,
            room_location=room_location,
            installment_date=installment_date if installment_date else None,
            purchase_date=purchase_date if purchase_date else None,
            warranty_expire_date=warranty_expire_date if warranty_expire_date else None,
            installation_notes=installation_notes,
            maintenance_period_days=m_days,
            frequency_unit=frequency_unit,
            last_maintenance_date=last_maintenance_date if last_maintenance_date else None,
            next_service_due=next_service_due if next_service_due else None,
            maintenance_notes=maintenance_notes,
            status=status if status else 'active',
            custom_fields=custom_fields,
            created_by=request.user,
            is_active=True,
            deleted=False,
        )
        obj.save()

        detail_url = reverse('backend:equipment_detail', kwargs={'data_id': obj.id})
        redirect_url = (
            reverse('backend:building_details', kwargs={'data_id': from_building_id})
            if from_building_id
            else reverse('backend:equipment_list')
        )
        building_name = obj.building.name if obj.building else "Unassigned"

        if is_ajax:
            messages.success(request, 'Equipment registered successfully.')
            return JsonResponse({
                'success': True,
                'equipment_id': obj.id,
                'equipment_code': obj.equipment_id,
                'equipment_name': obj.equipment_name,
                'building_name': building_name,
                'detail_url': detail_url,
                'redirect_url': redirect_url,
                'message': 'Equipment registered successfully.'
            })

        messages.success(request, 'Equipment registered successfully.')
        if from_building_id:
            return redirect('backend:building_details', data_id=from_building_id)
        return redirect('backend:equipment_list')

    # Initial type from GET if provided
    initial_type_id = request.GET.get('equipment_type_id') or request.GET.get('type')

    context = {
        'equipment_types': equipment_types,
        'equipment_types_json': json.dumps(equipment_types_data),
        'buildings': buildings,
        'buildings_json': json.dumps(buildings_data),
        'brands': all_brands,
        'from_building': from_building_id,
        'initial_type_id': initial_type_id,
    }
    return render(request, 'equipment/add.html', context)



@login_required
def equipment_update(request, data_id):
    if not checkUserPermission(request, 'can_edit', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to edit this equipment.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Equipment.objects.select_related('equipment_type', 'building'), pk=data_id, deleted=False)

    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True).order_by('name')
    buildings = Building.objects.filter(deleted=False, is_active=True).select_related('division', 'subdivision', 'section').order_by('name')

    existing_brands = list(
        Equipment.objects.filter(deleted=False)
        .exclude(brand='')
        .values_list('brand', flat=True)
        .distinct()
    )
    default_brands = ['Daikin', 'Carrier', 'Otis', 'KONE', 'Schindler', 'Caterpillar', 'Cummins', 'Mitsubishi', 'Siemens', 'Honeywell', 'Schneider Electric', 'ABB', 'Trane', 'York', 'Grundfos']
    all_brands = sorted(list(set(existing_brands + default_brands)))

    equipment_types_data = []
    for et in equipment_types:
        fields = [
            {
                'id': f.id,
                'field_name': f.field_name,
                'field_type': f.field_type,
                'is_required': f.is_required,
            }
            for f in et.data_fields.filter(is_active=True, deleted=False).order_by('id')
        ]
        equipment_types_data.append({
            'id': et.id,
            'name': et.name,
            'description': et.description or '',
            'fields': fields,
        })

    buildings_data = []
    for b in buildings:
        sec = b.section.name if b.section else ""
        sub = b.subdivision.name if b.subdivision else ""
        div = b.division.name if b.division else ""
        org_parts = [p for p in [sec, sub, div] if p]
        org_subdivision = " / ".join(org_parts) if org_parts else ""
        buildings_data.append({
            'id': b.id,
            'name': b.name,
            'code': b.building_id or '',
            'org_subdivision': org_subdivision,
        })

    if request.method == 'POST':
        is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.POST.get('is_ajax') == '1'

        equipment_id = request.POST.get('equipment_id', '').strip()
        equipment_name = request.POST.get('equipment_name', '').strip()
        equipment_type_id = request.POST.get('equipment_type_id') or None
        serial_number = request.POST.get('serial_number', '').strip()
        model_number = request.POST.get('model_number', '').strip()
        brand = request.POST.get('brand', '').strip()

        custom_fields_raw = request.POST.get('custom_fields', '{}')
        try:
            custom_fields = json.loads(custom_fields_raw) if isinstance(custom_fields_raw, str) else custom_fields_raw
        except Exception:
            custom_fields = {}

        assign_to_building = request.POST.get('assign_to_building') in ['true', 'on', '1', True]
        building_id = request.POST.get('building_id') if assign_to_building else None
        floor_location = request.POST.get('floor_location', '').strip() if assign_to_building else ''
        room_location = request.POST.get('room_location', '').strip() if assign_to_building else ''

        installment_date_str = request.POST.get('installment_date', '').strip() if assign_to_building else ''
        purchase_date_str = request.POST.get('purchase_date', '').strip()
        warranty_expire_date_str = request.POST.get('warranty_expire_date', '').strip()
        installation_notes = request.POST.get('installation_notes', '').strip() if assign_to_building else ''

        maintenance_period_days_str = request.POST.get('maintenance_period_days', '90').strip()
        frequency_unit = request.POST.get('frequency_unit', 'Days').strip() or 'Days'
        last_maintenance_date_str = request.POST.get('last_maintenance_date', '').strip()
        next_service_due_str = request.POST.get('next_service_due', '').strip()
        maintenance_notes = request.POST.get('maintenance_notes', '').strip()
        status = request.POST.get('status', obj.status or 'active').strip() or 'active'

        errors = []
        if not equipment_type_id:
            errors.append('Equipment Type is required.')
        if not equipment_id:
            errors.append('Equipment ID is required.')
        elif Equipment.objects.filter(equipment_id=equipment_id, deleted=False).exclude(pk=obj.pk).exists():
            errors.append(f"Equipment ID '{equipment_id}' already exists.")
        if not equipment_name:
            errors.append('Equipment Name is required.')
        if not serial_number:
            errors.append('Serial Number is required.')
        if not model_number:
            errors.append('Model Number is required.')
        if not brand:
            errors.append('Brand is required.')

        if assign_to_building:
            if not building_id:
                errors.append('Building is required when assigning to a building.')
            if not installment_date_str:
                errors.append('Installation Date is required when assigning to a building.')

        try:
            maintenance_period_days = int(maintenance_period_days_str) if maintenance_period_days_str else 90
            if maintenance_period_days <= 0:
                maintenance_period_days = 90
        except ValueError:
            maintenance_period_days = 90

        if errors:
            if is_ajax:
                return JsonResponse({'success': False, 'errors': errors}, status=400)
            for err in errors:
                messages.error(request, err)
            return redirect('backend:equipment_update', data_id=obj.id)

        installment_date = None
        if installment_date_str:
            try:
                installment_date = timezone.make_aware(datetime.strptime(installment_date_str, '%Y-%m-%d'))
            except ValueError:
                installment_date = None

        purchase_date = None
        if purchase_date_str:
            try:
                purchase_date = datetime.strptime(purchase_date_str, '%Y-%m-%d').date()
            except ValueError:
                purchase_date = None

        warranty_expire_date = None
        if warranty_expire_date_str:
            try:
                warranty_expire_date = datetime.strptime(warranty_expire_date_str, '%Y-%m-%d').date()
            except ValueError:
                warranty_expire_date = None

        last_maintenance_date = None
        if last_maintenance_date_str:
            try:
                last_maintenance_date = datetime.strptime(last_maintenance_date_str, '%Y-%m-%d').date()
            except ValueError:
                last_maintenance_date = None

        next_service_due = None
        if next_service_due_str:
            try:
                next_service_due = datetime.strptime(next_service_due_str, '%Y-%m-%d').date()
            except ValueError:
                next_service_due = None
        elif last_maintenance_date:
            if frequency_unit == 'Months':
                next_service_due = last_maintenance_date + timedelta(days=maintenance_period_days * 30)
            elif frequency_unit == 'Years':
                next_service_due = last_maintenance_date + timedelta(days=maintenance_period_days * 365)
            else:
                next_service_due = last_maintenance_date + timedelta(days=maintenance_period_days)

        obj.equipment_id = equipment_id
        obj.equipment_name = equipment_name
        obj.equipment_type_id = int(equipment_type_id)
        obj.building_id = int(building_id) if building_id else None
        obj.brand = brand
        obj.serial_number = serial_number
        obj.model_number = model_number
        obj.floor_location = floor_location
        obj.room_location = room_location
        obj.installment_date = installment_date
        obj.purchase_date = purchase_date
        obj.warranty_expire_date = warranty_expire_date
        obj.installation_notes = installation_notes
        obj.maintenance_period_days = maintenance_period_days
        obj.frequency_unit = frequency_unit
        obj.last_maintenance_date = last_maintenance_date
        obj.next_service_due = next_service_due
        obj.maintenance_notes = maintenance_notes
        obj.status = status
        obj.custom_fields = custom_fields
        obj.updated_by = request.user
        obj.save()

        redirect_url = reverse('backend:equipment_list')
        if is_ajax:
            messages.success(request, 'Equipment updated successfully.')
            return JsonResponse({
                'success': True,
                'equipment_id': obj.id,
                'equipment_code': obj.equipment_id,
                'equipment_name': obj.equipment_name,
                'building_name': obj.building.name if obj.building else 'Unassigned',
                'detail_url': reverse('backend:equipment_detail', kwargs={'data_id': obj.id}),
                'redirect_url': redirect_url,
                'message': 'Equipment updated successfully.',
            })

        messages.success(request, 'Equipment updated successfully.')
        return redirect('backend:equipment_list')

    context = {
        'is_edit': True,
        'obj': obj,
        'equipment_types': equipment_types,
        'equipment_types_json': json.dumps(equipment_types_data),
        'buildings': buildings,
        'buildings_json': json.dumps(buildings_data),
        'brands': all_brands,
        'from_building': obj.building_id,
        'initial_type_id': obj.equipment_type_id,
        'existing_custom_fields_json': json.dumps(obj.custom_fields or {}),
    }
    return render(request, 'equipment/add.html', context)


@login_required
def equipment_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to change status of this equipment.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        obj = get_object_or_404(Equipment, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Equipment status changed to {status_text} successfully.')
        return redirect('backend:equipment_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:equipment_list')


@login_required
def equipment_detail(request, data_id):
    if not checkUserPermission(request, 'can_view', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to view this equipment.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Equipment, pk=data_id, deleted=False)
    
    # Get active components
    components = obj.components.filter(deleted=False).order_by('name')
    
    context = {
        'obj': obj,
        'components': components,
    }
    return render(request, 'equipment/detail.html', context)




def generate_qr_code(data):
    """
    Generate QR code for the given data.
    """
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    return img 

import base64
from io import BytesIO

def qr_equipement_views(request, data_id):
    obj = get_object_or_404(Equipment, pk=data_id, deleted=False)
    building_param = obj.building_id or ''
    ticket_param = (request.GET.get('ticket') or '').strip()
    technician_param = (request.GET.get('technician') or '').strip()

    extra_qs = f"&ticket={ticket_param}" if ticket_param else ""
    if technician_param:
        extra_qs += f"&technician={technician_param}"

    if request.GET.get('view') == '1':
        last_maintenance = obj.maintenance_records.filter(deleted=False).order_by('-maintenance_date').first()
        ticket_obj = None
        technician_obj = None
        if ticket_param.isdigit():
            ticket_obj = Ticket.objects.filter(pk=int(ticket_param), deleted=False).select_related('building', 'equipment').first()
        if technician_param.isdigit():
            technician_obj = Technician.objects.filter(pk=int(technician_param), deleted=False).first()

        add_url = request.build_absolute_uri(reverse('backend:static_otp_verification'))
        qr_data = f"{add_url}?building={building_param}&equipment={obj.id}{extra_qs}"
        img = generate_qr_code(qr_data)
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        qr_base64 = base64.b64encode(buffer.getvalue()).decode()

        context = {
            'obj': obj,
            'last_maintenance': last_maintenance,
            'qr_base64': qr_base64,
            'ticket_obj': ticket_obj,
            'technician_obj': technician_obj,
            'ticket_param': ticket_param,
            'technician_param': technician_param,
        }
        return render(request, "equipment/qr.html", context)

    # Redirect to static_otp_verification with auto-fill parameters
    redirect_url = reverse('backend:static_otp_verification')
    return redirect(f"{redirect_url}?building={building_param}&equipment={obj.id}{extra_qs}")

# ======================================== Maintenance Record Review & Ticket Workflow Helpers ========================================

def _recalculate_equipment_service_dates(equipment, service_date):
    """Update equipment's last_maintenance_date, next_service_due, and restore active status after approved service."""
    if not equipment or not service_date:
        return
    equipment.last_maintenance_date = service_date
    period = equipment.maintenance_period_days or 90
    unit = (equipment.frequency_unit or 'Days').strip()
    if unit == 'Weeks':
        delta = timedelta(days=period * 7)
    elif unit == 'Months':
        delta = timedelta(days=period * 30)
    elif unit == 'Years':
        delta = timedelta(days=period * 365)
    else:
        delta = timedelta(days=period)
    equipment.next_service_due = service_date + delta
    if equipment.status in ('in_repair', 'under_maintenance'):
        equipment.status = 'active'
    equipment.save()


def _approve_ticket_record(record, reviewer, review_notes=''):
    """Approve a submitted MaintenanceRecord, finalize its ticket, update equipment schedule, and log into Service History."""
    now = timezone.now()
    record.review_status = 'approved'
    record.maintainance_status = 'completed'
    record.review_notes = review_notes
    record.reviewed_by = reviewer
    record.reviewed_at = now
    record.updated_by = reviewer
    record.save()

    if record.equipment:
        _recalculate_equipment_service_dates(record.equipment, record.maintenance_date or timezone.localdate())

    if record.ticket:
        tkt = record.ticket
        tkt.status = 'resolved'
        tkt.resolved_at = now
        if record.duration_hours is not None:
            tkt.actual_hours = record.duration_hours
        tkt.updated_by = reviewer
        tkt.save()

        AssignActivities.objects.filter(ticket=tkt, is_active=True, status='accepted').update(
            completion_status='completed',
            completed_at=now,
            updated_by=reviewer,
        )

        reviewer_name = reviewer.get_full_name() or reviewer.username
        review_round = MaintenanceRecord.objects.filter(
            ticket=tkt, deleted=False, review_status__in=['approved', 'rejected']
        ).count()
        revision_count = MaintenanceRecord.objects.filter(
            ticket=tkt, deleted=False, review_status='rejected'
        ).count()
        rev_suffix = f" after {revision_count} revision{'s' if revision_count != 1 else ''}" if revision_count else ""
        note_suffix = f" Review notes: {review_notes}" if review_notes else ""
        TicketActivityLog.objects.create(
            ticket=tkt,
            actor=reviewer,
            action_type='review_approved',
            description=(
                f"Review #{review_round} (Approved) — Supervisor ({reviewer_name}) approved SR-{record.id:03d}"
                f"{rev_suffix}. Ticket marked Resolved and moved to Service History.{note_suffix}"
            ),
            metadata={
                'record_id': record.id,
                'review_status': 'approved',
                'review_round': review_round,
                'revision_count': revision_count,
            },
        )


def _reject_ticket_record(record, reviewer, review_notes):
    """Reject a submitted MaintenanceRecord and return the ticket to In Progress for rework."""
    now = timezone.now()
    record.review_status = 'rejected'
    record.review_notes = review_notes
    record.reviewed_by = reviewer
    record.reviewed_at = now
    record.updated_by = reviewer
    record.save()

    if record.ticket:
        tkt = record.ticket
        tkt.status = 'in_progress'
        tkt.updated_by = reviewer
        tkt.save()

        AssignActivities.objects.filter(ticket=tkt, is_active=True, status='accepted').update(
            completion_status='in_progress',
            updated_by=reviewer,
        )

        reviewer_name = reviewer.get_full_name() or reviewer.username
        review_round = MaintenanceRecord.objects.filter(
            ticket=tkt, deleted=False, review_status__in=['approved', 'rejected']
        ).count()
        revision_count = MaintenanceRecord.objects.filter(
            ticket=tkt, deleted=False, review_status='rejected'
        ).count()
        TicketActivityLog.objects.create(
            ticket=tkt,
            actor=reviewer,
            action_type='review_rejected',
            description=(
                f"Review #{review_round} (Revision #{revision_count} Requested) — Supervisor ({reviewer_name}) "
                f"sent SR-{record.id:03d} back to technician for rework. Required changes: {review_notes}"
            ),
            metadata={
                'record_id': record.id,
                'review_status': 'rejected',
                'review_round': review_round,
                'revision_count': revision_count,
            },
        )


@login_required
def maintenance_record_review(request, record_id):
    """Approve or Reject a MaintenanceRecord from detail/review pages."""
    record = get_object_or_404(MaintenanceRecord, pk=record_id, deleted=False)
    referer = request.META.get('HTTP_REFERER')
    fallback_url = (
        reverse('backend:ticket_detail', kwargs={'data_id': record.ticket_id})
        if record.ticket_id
        else reverse('backend:equipment_detail', kwargs={'data_id': record.equipment_id})
    )
    redirect_url = referer if referer else fallback_url

    if request.method != 'POST':
        messages.error(request, 'Invalid request method.')
        return redirect(redirect_url)

    action = request.POST.get('action', '').strip()   # 'approve' or 'reject'
    review_notes = request.POST.get('review_notes', '').strip()

    if action == 'approve':
        _approve_ticket_record(record, request.user, review_notes)
        messages.success(request, 'Maintenance record approved and logged to Service History.')

    elif action == 'reject':
        if not review_notes:
            messages.error(request, 'Please provide a reason for rejection.')
            return redirect(redirect_url)
        _reject_ticket_record(record, request.user, review_notes)
        messages.success(request, 'Revision requested; maintenance record rejected.')

    else:
        messages.error(request, 'Unknown action.')

    return redirect(redirect_url)

@login_required
def equipment_history(request):
    if not checkUserPermission(request, 'can_view', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)

    # ── Filters ──────────────────────────────────────────────────────────
    search        = request.GET.get('search', '').strip()
    building_id   = request.GET.get('building_id', '').strip()
    record_type   = request.GET.get('record_type', '').strip()
    technician_id = (request.GET.get('technician_id') or request.GET.get('technician') or '').strip()
    source_filter = request.GET.get('source', '').strip()
    date_filter   = request.GET.get('date', '').strip()

    # Ticket-sourced records only enter Service History after supervisor review approval
    qs = (
        MaintenanceRecord.objects.filter(deleted=False)
        .exclude(ticket__isnull=False, review_status__in=['pending', 'rejected'])
        .select_related(
            'equipment', 'equipment__building', 'equipment__equipment_type',
            'maintenance', 'ticket', 'technician', 'assigned_to', 'reviewed_by'
        )
        .order_by('-maintenance_date', '-created_at')
    )

    if search:
        qs = qs.filter(
            Q(equipment__equipment_id__icontains=search)
            | Q(equipment__equipment_name__icontains=search)
            | Q(equipment__building__name__icontains=search)
            | Q(maintenance__maintenance_serial__icontains=search)
            | Q(work_description__icontains=search)
            | Q(technician__first_name__icontains=search)
            | Q(technician__last_name__icontains=search)
            | Q(technician__username__icontains=search)
            | Q(assigned_to__first_name__icontains=search)
            | Q(assigned_to__last_name__icontains=search)
        )
    if building_id:
        qs = qs.filter(equipment__building_id=building_id)
    if record_type:
        qs = qs.filter(record_type=record_type)
    if technician_id:
        qs = qs.filter(
            Q(assigned_to_id=technician_id)
            | Q(technician__technician_profile__id=technician_id)
        )
    if source_filter == 'Scheduled':
        qs = qs.filter(maintenance__isnull=False)
    elif source_filter == 'Request':
        qs = qs.filter(ticket__isnull=False)
    elif source_filter == 'Manual':
        qs = qs.filter(maintenance__isnull=True, ticket__isnull=True)
    if date_filter:
        try:
            qs = qs.filter(maintenance_date=datetime.strptime(date_filter, '%Y-%m-%d').date())
        except ValueError:
            pass

    # ── Pagination ────────────────────────────────────────────────────────
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count

    # ── Filter options ────────────────────────────────────────────────────
    buildings = Building.objects.filter(deleted=False, is_active=True).order_by('name')
    technicians = Technician.objects.filter(deleted=False, is_active=True).order_by('first_name', 'last_name')

    context = {
        'data_list'     : data_list,
        'paginator_list': paginator_list,
        'last_page'     : last_page,
        'per_page'      : per_page,
        'filtered_total_count': filtered_total_count,
        # active filters
        'search'        : search,
        'f_building_id' : building_id,
        'f_record_type' : record_type,
        'f_technician_id': technician_id,
        'f_source'      : source_filter,
        'f_date'        : date_filter,
        # filter options
        'buildings'     : buildings,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'technicians'   : technicians,
        'source_choices': ['Scheduled', 'Request', 'Manual'],
    }
    return render(request, 'equipment/history.html', context)


# ======================================== Ticket ========================================
# Ticket Management
# ======================================== 

@login_required
def ticket_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to view this ticket.')
        return render(request, '403.html', status=403)
    
    search = request.GET.get('search', '').strip()
    status_filter = request.GET.get('status', '').strip()
    priority_filter = request.GET.get('priority', '').strip()
    
    qs = Ticket.objects.filter(deleted=False).select_related('equipment', 'building', 'user').order_by('-created_at')
    
    if search:
        qs = qs.filter(
            Q(title__icontains=search)
            | Q(equipment__equipment_id__icontains=search)
            | Q(issue_description__icontains=search)
            | Q(building__name__icontains=search)
        )
    if status_filter:
        qs = qs.filter(status=status_filter)
    if priority_filter:
        qs = qs.filter(priority=priority_filter)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count
    
    context = {
        'data_list': data_list,
        'tickets': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'status_filter': status_filter,
        'priority_filter': priority_filter,
        'status_choices': Ticket.STATUS_CHOICES,
        'priority_choices': Ticket.PRIORITY_CHOICES,
    }
    return render(request, 'ticket/list.html', context)

@login_required
def ticket_detail(request, data_id):
    if not checkUserPermission(request, 'can_view', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to view this ticket.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(
        Ticket.objects.select_related(
            'equipment', 'equipment__equipment_type', 'building', 'user', 'created_by'
        ),
        pk=data_id,
        deleted=False,
    )

    def _safe_redirect():
        next_url = (request.POST.get('next') or '').strip()
        if next_url.startswith('/') and not next_url.startswith('//'):
            return redirect(next_url)
        return redirect('backend:ticket_detail', data_id=obj.id)

    def _get_or_create_activity(tech_obj):
        act = AssignActivities.objects.filter(
            ticket=obj, technician=tech_obj, is_active=True
        ).order_by('-updated_at', '-id').first()
        if not act:
            act = AssignActivities.objects.create(
                ticket=obj,
                technician=tech_obj,
                status='pending',
                completion_status='pending',
                created_by=request.user,
            )
        return act

    if request.method == 'POST':
        if not checkUserPermission(request, 'can_update', '/backend/ticket/ticket/'):
            messages.error(request, 'You do not have permission to update this ticket.')
            return _safe_redirect()

        action = (request.POST.get('action') or '').strip()
        now = timezone.now()

        # ── 1. Technician Accepts Assigned Task ─────────────────────────────
        if action == 'accept_task':
            tech_id = (request.POST.get('technician_id') or '').strip()
            assign_qs = obj.assigned_technicians.filter(deleted=False).select_related('technician')
            if tech_id:
                assign_link = assign_qs.filter(technician_id=tech_id).first()
            else:
                # Prefer first pending assignment
                assign_link = None
                for al in assign_qs:
                    act_check = AssignActivities.objects.filter(
                        ticket=obj, technician=al.technician, is_active=True
                    ).first()
                    if not act_check or act_check.status == 'pending':
                        assign_link = al
                        break
                if not assign_link:
                    assign_link = assign_qs.first()

            if not assign_link or not assign_link.technician:
                messages.error(request, 'No assigned technician found to accept this task.')
                return _safe_redirect()

            tech_obj = assign_link.technician
            act = _get_or_create_activity(tech_obj)
            if act.status == 'accepted':
                messages.info(request, f'{tech_obj.display_name} has already accepted this task.')
                return _safe_redirect()

            act.status = 'accepted'
            act.accepted_at = now
            if act.completion_status in ('cancelled', 'skipped'):
                act.completion_status = 'pending'
            act.updated_by = request.user
            act.save()

            if obj.status in ('open', 'assigned'):
                obj.status = 'accepted'
                obj.updated_by = request.user
                obj.save()

            TicketActivityLog.objects.create(
                ticket=obj,
                actor=request.user,
                action_type='accepted',
                description=f"{tech_obj.display_name} accepted the assigned task and can now start work.",
                metadata={'technician_id': tech_obj.id, 'technician_name': tech_obj.display_name},
            )
            messages.success(request, f'Task accepted by {tech_obj.display_name}. You can now start work.')
            return _safe_redirect()

        # ── 2. Technician Rejects Assigned Task ─────────────────────────────
        if action == 'reject_task':
            tech_id = (request.POST.get('technician_id') or '').strip()
            reject_reason = (request.POST.get('reject_reason') or '').strip()
            assign_qs = obj.assigned_technicians.filter(deleted=False).select_related('technician')
            assign_link = assign_qs.filter(technician_id=tech_id).first() if tech_id else assign_qs.first()

            if not assign_link or not assign_link.technician:
                messages.error(request, 'No assigned technician found for this action.')
                return _safe_redirect()

            tech_obj = assign_link.technician
            act = _get_or_create_activity(tech_obj)
            act.status = 'rejected'
            act.completion_status = 'cancelled'
            act.updated_by = request.user
            act.save()

            reason_txt = f" Reason: {reject_reason}" if reject_reason else ""
            TicketActivityLog.objects.create(
                ticket=obj,
                actor=request.user,
                action_type='rejected_assignment',
                description=f"{tech_obj.display_name} declined the task assignment.{reason_txt}",
                metadata={'technician_id': tech_obj.id, 'reason': reject_reason},
            )

            # If no active pending/accepted assignments remain, reopen ticket for reassignment
            still_active = AssignActivities.objects.filter(
                ticket=obj, is_active=True, status__in=['pending', 'accepted']
            ).exists()
            if not still_active and obj.status not in ('resolved', 'closed'):
                obj.status = 'open'
                obj.updated_by = request.user
                obj.save()

            messages.warning(request, f'Assignment declined by {tech_obj.display_name}.')
            return _safe_redirect()

        # ── 3. Technician Starts Work (STRICT: Must Accept First) ───────────
        if action == 'start_work':
            tech_id = (request.POST.get('technician_id') or '').strip()
            accepted_qs = AssignActivities.objects.filter(
                ticket=obj, is_active=True, status='accepted'
            ).select_related('technician')
            if tech_id:
                accepted_act = accepted_qs.filter(technician_id=tech_id).first()
            else:
                accepted_act = accepted_qs.first()

            if not accepted_act:
                messages.error(request, 'Technician must accept the task first before starting work.')
                return _safe_redirect()

            accepted_act.completion_status = 'in_progress'
            if not accepted_act.started_at:
                accepted_act.started_at = now
            accepted_act.updated_by = request.user
            accepted_act.save()

            obj.status = 'in_progress'
            obj.updated_by = request.user
            obj.save()

            if obj.equipment and obj.equipment.status == 'active':
                obj.equipment.status = 'under_maintenance'
                obj.equipment.save()

            tech_obj = accepted_act.technician
            TicketActivityLog.objects.create(
                ticket=obj,
                actor=request.user,
                action_type='work_started',
                description=f"{tech_obj.display_name} started maintenance work on {obj.equipment.equipment_name or obj.equipment.equipment_id if obj.equipment else 'equipment'}.",
                metadata={'technician_id': tech_obj.id, 'technician_name': tech_obj.display_name},
            )
            messages.success(request, f'Work started by {tech_obj.display_name}. Ticket is now In Progress.')
            return _safe_redirect()

        # ── 4. Technician Submits Completed Work for Supervisor Review ──────
        if action in ('submit_work', 'complete_service'):
            if not obj.equipment:
                messages.error(request, 'Cannot submit work: this ticket has no linked equipment.')
                return _safe_redirect()

            accepted_acts = AssignActivities.objects.filter(
                ticket=obj, is_active=True, status='accepted'
            ).select_related('technician')
            if not accepted_acts.exists():
                messages.error(request, 'An assigned technician must accept and start work before submitting.')
                return _safe_redirect()

            has_started = (
                obj.status == 'in_progress'
                or accepted_acts.filter(completion_status='in_progress').exists()
            )
            if not has_started:
                messages.error(request, 'Please start work on the task before submitting completed work for review.')
                return _safe_redirect()

            work_description = request.POST.get('work_description', '').strip() or obj.issue_description or obj.title
            record_type = request.POST.get('record_type', 'repair').strip() or 'repair'
            maint_date_str = request.POST.get('maintenance_date', '').strip()
            parts_replaced = request.POST.get('parts_replaced', '').strip()
            cost_raw = request.POST.get('cost', '').strip()
            duration_raw = request.POST.get('duration_hours', '').strip()
            service_tech_id = request.POST.get('service_technician_id', '').strip()

            try:
                maint_date = datetime.strptime(maint_date_str, '%Y-%m-%d').date() if maint_date_str else timezone.localdate()
            except ValueError:
                maint_date = timezone.localdate()

            assigned_tech = None
            tech_user = None
            if service_tech_id:
                assigned_tech = Technician.objects.filter(id=service_tech_id, deleted=False).first()
                if assigned_tech:
                    tech_user = assigned_tech.user
            if not assigned_tech:
                first_accepted = accepted_acts.first()
                if first_accepted:
                    assigned_tech = first_accepted.technician
                    tech_user = assigned_tech.user

            cost_val = None
            if cost_raw:
                try:
                    cost_val = float(cost_raw)
                except ValueError:
                    cost_val = None

            duration_val = None
            if duration_raw:
                try:
                    duration_val = float(duration_raw)
                except ValueError:
                    duration_val = None

            record = MaintenanceRecord.objects.create(
                equipment=obj.equipment,
                ticket=obj,
                record_type=record_type,
                maintenance_date=maint_date,
                assigned_to=assigned_tech,
                technician=tech_user or request.user,
                work_description=work_description,
                parts_replaced=parts_replaced or '',
                cost=cost_val,
                duration_hours=duration_val,
                maintainance_status='completed',
                review_status='pending',
                created_by=request.user,
            )

            obj.status = 'pending_review'
            obj.updated_by = request.user
            obj.save()

            accepted_acts.update(
                completion_status='submitted',
                updated_by=request.user,
            )

            tech_label = assigned_tech.display_name if assigned_tech else (request.user.get_full_name() or request.user.username)
            submission_num = MaintenanceRecord.objects.filter(ticket=obj, deleted=False).count()
            prior_revisions = MaintenanceRecord.objects.filter(
                ticket=obj, deleted=False, review_status='rejected'
            ).count()
            if submission_num > 1 and prior_revisions > 0:
                sub_prefix = f"Submission #{submission_num} (Updated after Revision #{prior_revisions}) — "
                msg_text = f'Updated work (Submission #{submission_num}) submitted for supervisor review.'
            else:
                sub_prefix = f"Submission #{submission_num} — "
                msg_text = 'Work submitted for supervisor review. Once approved, it will move to Service History.'

            TicketActivityLog.objects.create(
                ticket=obj,
                actor=request.user,
                action_type='submitted_for_review',
                description=(
                    f"{sub_prefix}{tech_label} submitted work for supervisor review "
                    f"(SR-{record.id:03d}): {work_description}"
                ),
                metadata={
                    'record_id': record.id,
                    'technician_id': assigned_tech.id if assigned_tech else None,
                    'submission_num': submission_num,
                    'prior_revisions': prior_revisions,
                },
            )
            messages.success(request, msg_text)
            return _safe_redirect()

        # ── 5. Supervisor Approves Review → Moves Ticket to Service History ─
        if action == 'approve_review':
            record_id = (request.POST.get('record_id') or '').strip()
            review_notes = (request.POST.get('review_notes') or '').strip()
            rec_qs = MaintenanceRecord.objects.filter(ticket=obj, deleted=False)
            if record_id:
                record = rec_qs.filter(pk=record_id).first()
            else:
                record = rec_qs.filter(review_status='pending').order_by('-created_at').first()

            if not record:
                messages.error(request, 'No pending work submission found to approve.')
                return _safe_redirect()

            _approve_ticket_record(record, request.user, review_notes)
            messages.success(request, f'Work approved! Ticket #{obj.id} is resolved and moved to Service History (SR-{record.id:03d}).')
            return _safe_redirect()

        # ── 6. Supervisor Rejects Review → Back to In Progress (Rework) ─────
        if action == 'reject_review':
            record_id = (request.POST.get('record_id') or '').strip()
            review_notes = (request.POST.get('review_notes') or '').strip()
            if not review_notes:
                messages.error(request, 'Please provide revision notes explaining why the submission was rejected.')
                return _safe_redirect()

            rec_qs = MaintenanceRecord.objects.filter(ticket=obj, deleted=False)
            if record_id:
                record = rec_qs.filter(pk=record_id).first()
            else:
                record = rec_qs.filter(review_status='pending').order_by('-created_at').first()

            if not record:
                messages.error(request, 'No pending work submission found to review.')
                return _safe_redirect()

            _reject_ticket_record(record, request.user, review_notes)
            messages.warning(request, 'Revision requested. Ticket returned to In Progress for rework.')
            return _safe_redirect()

        # ── 7. Handle Comment Submission ────────────────────────────────────
        comment_body = request.POST.get('comment_body', '').strip()
        if comment_body:
            TicketComment.objects.create(
                ticket=obj,
                author=request.user,
                body=comment_body,
                created_by=request.user,
            )
            TicketActivityLog.objects.create(
                ticket=obj,
                actor=request.user,
                action_type='comment_added',
                description=comment_body,
            )
            messages.success(request, 'Comment added successfully.')
            return _safe_redirect()

        # ── 8. Handle Technician Assignment ─────────────────────────────────
        technician_id = request.POST.get('technician_id')
        if technician_id and (action in ('', 'assign_technician')):
            try:
                technician_obj = Technician.objects.get(id=technician_id, deleted=False)

                existing_link = AssignTechnician.objects.filter(
                    ticket=obj, technician=technician_obj, deleted=False
                ).first()
                existing_act = AssignActivities.objects.filter(
                    ticket=obj, technician=technician_obj, is_active=True
                ).first()

                if existing_link and existing_act and existing_act.status != 'rejected':
                    messages.warning(request, f'{technician_obj.display_name} is already assigned to this ticket.')
                else:
                    if not existing_link:
                        AssignTechnician.objects.create(
                            ticket=obj,
                            technician=technician_obj,
                            assign_by=request.user,
                        )
                    if existing_act:
                        existing_act.status = 'pending'
                        existing_act.completion_status = 'pending'
                        existing_act.accepted_at = None
                        existing_act.started_at = None
                        existing_act.completed_at = None
                        existing_act.updated_by = request.user
                        existing_act.save()
                        log_verb = "Re-assigned"
                    else:
                        AssignActivities.objects.create(
                            ticket=obj,
                            technician=technician_obj,
                            status='pending',
                            completion_status='pending',
                            created_by=request.user,
                        )
                        log_verb = "Assigned"

                    if obj.status in ('open', 'rejected'):
                        obj.status = 'assigned'
                        obj.assigned_at = now
                        obj.updated_by = request.user
                        obj.save()

                    TicketActivityLog.objects.create(
                        ticket=obj,
                        actor=request.user,
                        action_type='assigned',
                        description=f"{log_verb} technician {technician_obj.display_name} — waiting for technician to accept.",
                        metadata={'technician_id': technician_obj.id, 'technician_name': technician_obj.display_name},
                    )
                    messages.success(request, f'{technician_obj.display_name} assigned. Waiting for technician acceptance.')
            except Technician.DoesNotExist:
                messages.error(request, 'Selected technician does not exist.')

        return _safe_redirect()

    # ── GET: Build Rich Workflow Context ────────────────────────────────────
    technicians = Technician.objects.filter(deleted=False, is_active=True).order_by('first_name', 'last_name')
    service_records = list(
        MaintenanceRecord.objects.filter(ticket=obj, deleted=False)
        .select_related('assigned_to', 'technician', 'reviewed_by')
        .prefetch_related('components')
        .order_by('-created_at', '-id')
    )
    checklist_items = list(
        obj.checklist_items.filter(deleted=False, is_active=True).order_by('order', 'id')
    )
    checklist_total = len(checklist_items)
    checklist_done = sum(1 for i in checklist_items if i.is_completed)
    pending_record = next((r for r in service_records if r.review_status == 'pending'), None)
    approved_record = next((r for r in service_records if r.review_status == 'approved'), None)
    rejected_record = next((r for r in service_records if r.review_status == 'rejected'), None)

    submission_count = len(service_records)
    revision_count = sum(1 for r in service_records if r.review_status == 'rejected')
    review_count = sum(1 for r in service_records if r.review_status in ('approved', 'rejected'))
    next_review_round = review_count + 1

    # Number every submission chronologically (#1, #2, ...) for revision history
    chronological_records = list(reversed(service_records))
    numbered_rejected_records = []
    rev_counter = 0
    for idx, rec in enumerate(chronological_records, start=1):
        if rec.review_status == 'rejected':
            rev_counter += 1
            numbered_rejected_records.append({
                'record': rec,
                'attempt_num': idx,
                'revision_num': rev_counter,
            })
    numbered_rejected_records.reverse()

    raw_assignments = list(
        obj.assigned_technicians.filter(deleted=False)
        .select_related('technician', 'assign_by')
        .order_by('assigned_at')
    )
    act_map = {}
    for act in AssignActivities.objects.filter(ticket=obj, is_active=True).order_by('-updated_at', '-id'):
        if act.technician_id not in act_map:
            act_map[act.technician_id] = act

    assigned_rows = []
    for link in raw_assignments:
        if not link.technician:
            continue
        act = act_map.get(link.technician_id)
        raw_acc = (act.status if act else 'pending') or 'pending'
        acc_status = 'accepted' if raw_acc == 'completed' else raw_acc
        comp_status = (act.completion_status if act else 'pending') or 'pending'
        assigned_rows.append({
            'link': link,
            'technician': link.technician,
            'activity': act,
            'acceptance_status': acc_status,
            'completion_status': comp_status,
        })

    has_assigned = any(r['acceptance_status'] in ('pending', 'accepted') for r in assigned_rows)
    has_accepted = any(r['acceptance_status'] == 'accepted' for r in assigned_rows)
    has_pending_acceptance = any(r['acceptance_status'] == 'pending' for r in assigned_rows)
    first_accepted_row = next((r for r in assigned_rows if r['acceptance_status'] == 'accepted'), None)
    first_pending_row = next((r for r in assigned_rows if r['acceptance_status'] == 'pending'), None)
    default_tech_id = (
        first_accepted_row['technician'].id if first_accepted_row
        else (assigned_rows[0]['technician'].id if assigned_rows else '')
    )

    # Compute current stage (1..6) for the visual progress stepper
    if obj.status in ('resolved', 'closed') or approved_record:
        current_step = 6
        stage_key = 'resolved'
        stage_label = 'Approved & In Service History'
    elif obj.status == 'pending_review' or pending_record:
        current_step = 5
        stage_key = 'pending_review'
        stage_label = 'Pending Supervisor Review'
    elif obj.status == 'in_progress' or any(r['completion_status'] == 'in_progress' for r in assigned_rows):
        current_step = 4
        stage_key = 'in_progress'
        stage_label = 'Work In Progress'
    elif has_accepted or obj.status == 'accepted':
        current_step = 3
        stage_key = 'accepted'
        stage_label = 'Accepted by Technician'
    elif has_assigned or obj.status == 'assigned':
        current_step = 2
        stage_key = 'assigned'
        stage_label = 'Assigned (Awaiting Acceptance)'
    else:
        current_step = 1
        stage_key = 'open'
        stage_label = 'Request Received'

    step5_sub = (
        f"Review #{next_review_round}" if current_step == 5
        else (f"Approved ({review_count} review{'s' if review_count != 1 else ''})" if current_step == 6
              else (f"Revision #{revision_count} in rework" if revision_count > 0 else 'Pending work'))
    )
    workflow_steps = [
        {'num': 1, 'label': 'Request Received', 'sub': obj.building.name if obj.building else 'Building Request'},
        {'num': 2, 'label': 'Technician Assigned', 'sub': f"{len([r for r in assigned_rows if r['acceptance_status'] != 'rejected'])} assigned" if has_assigned else 'Unassigned'},
        {'num': 3, 'label': 'Task Accepted', 'sub': first_accepted_row['technician'].display_name if first_accepted_row else 'Awaiting acceptance'},
        {'num': 4, 'label': 'Work Started', 'sub': 'In progress' if current_step == 4 else ('Completed' if current_step > 4 else 'Locked until accepted')},
        {'num': 5, 'label': 'Submitted & Review', 'sub': step5_sub},
        {'num': 6, 'label': 'Service History', 'sub': f"SR-{approved_record.id:03d}" if approved_record else 'Final archive'},
    ]

    activity_logs = list(
        obj.activity_log.select_related('actor').order_by('created_at', 'id')
    )
    has_created_log = any(log.action_type == 'created' for log in activity_logs)

    comments = list(
        obj.comments.filter(deleted=False).select_related('author').order_by('created_at')
    )

    context = {
        'obj': obj,
        'technicians': technicians,
        'assigned_rows': assigned_rows,
        'has_assigned': has_assigned,
        'has_accepted': has_accepted,
        'has_pending_acceptance': has_pending_acceptance,
        'first_accepted_row': first_accepted_row,
        'first_pending_row': first_pending_row,
        'current_step': current_step,
        'stage_key': stage_key,
        'stage_label': stage_label,
        'workflow_steps': workflow_steps,
        'service_records': service_records,
        'pending_record': pending_record,
        'approved_record': approved_record,
        'rejected_record': rejected_record,
        'submission_count': submission_count,
        'revision_count': revision_count,
        'review_count': review_count,
        'next_review_round': next_review_round,
        'numbered_rejected_records': numbered_rejected_records,
        'checklist_items': checklist_items,
        'checklist_total': checklist_total,
        'checklist_done': checklist_done,
        'activity_logs': activity_logs,
        'has_created_log': has_created_log,
        'comments': comments,
        'default_tech_id': default_tech_id,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'today_iso': timezone.localdate().strftime('%Y-%m-%d'),
    }
    return render(request, 'ticket/detail.html', context)

@login_required
def ticket_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to add this ticket.')
        return render(request, '403.html', status=403)

    equipments = (
        Equipment.objects.filter(deleted=False, is_active=True)
        .select_related('equipment_type', 'building')
        .order_by('equipment_name', 'equipment_id')
    )
    buildings = Building.objects.filter(deleted=False, is_active=True).order_by('name')

    if request.method == 'POST':
        title = request.POST.get('title', '').strip()
        equipment_id = request.POST.get('equipment_id', '').strip()
        building_id = request.POST.get('building_id', '').strip()
        issue_type = request.POST.get('issue_type', 'breakdown').strip() or 'breakdown'
        priority = request.POST.get('priority', 'medium').strip() or 'medium'
        status = request.POST.get('status', 'open').strip() or 'open'
        issue_description = request.POST.get('issue_description', '').strip()
        due_date_raw = request.POST.get('due_date', '').strip()
        est_hours_raw = request.POST.get('estimated_hours', '').strip()

        due_date_val = None
        if due_date_raw:
            try:
                parsed_d = datetime.strptime(due_date_raw, '%Y-%m-%d')
                due_date_val = timezone.make_aware(parsed_d) if timezone.is_naive(parsed_d) else parsed_d
            except ValueError:
                due_date_val = None

        est_hours_val = None
        if est_hours_raw:
            try:
                est_hours_val = float(est_hours_raw)
            except ValueError:
                est_hours_val = None

        if not title or not equipment_id or not building_id:
            messages.error(request, 'Title, Equipment, and Building are required.')
        else:
            obj = Ticket(
                title=title,
                equipment_id=equipment_id,
                building_id=building_id,
                issue_type=issue_type,
                priority=priority,
                status=status,
                issue_description=issue_description,
                due_date=due_date_val,
                estimated_hours=est_hours_val,
                user=request.user,
                created_by=request.user,
            )
            obj.save()

            # Save checklist items added during ticket creation
            checklist_titles = [t.strip() for t in request.POST.getlist('checklist_items') if t.strip()]
            for idx, item_title in enumerate(checklist_titles, start=1):
                TicketTaskItem.objects.create(
                    ticket=obj,
                    title=item_title[:255],
                    order=idx,
                    created_by=request.user,
                )

            eq_label = (obj.equipment.equipment_name or obj.equipment.equipment_id) if obj.equipment else 'N/A'
            bld_label = obj.building.name if obj.building else 'N/A'
            desc_parts = [
                f"Maintenance request received from {bld_label} for {eq_label} "
                f"({obj.get_issue_type_display()}, {obj.get_priority_display()} priority)."
            ]
            if checklist_titles:
                desc_parts.append(f"Checklist ({len(checklist_titles)} task{'s' if len(checklist_titles) != 1 else ''}): " + "; ".join(checklist_titles) + ".")
            if obj.issue_description:
                desc_parts.append(f"Requested work: {obj.issue_description}")
            TicketActivityLog.objects.create(
                ticket=obj,
                actor=request.user,
                action_type='created',
                description=" ".join(desc_parts),
                metadata={'building_id': obj.building_id, 'equipment_id': obj.equipment_id, 'checklist_count': len(checklist_titles)},
            )
            messages.success(request, 'Ticket created successfully.')
            return redirect('backend:ticket_detail', data_id=obj.id)

    # Pre-fill equipment and building from GET parameters if provided
    initial_eq = (request.GET.get('equipment_id') or request.GET.get('equipment') or '').strip()
    initial_b = (request.GET.get('building_id') or request.GET.get('building') or '').strip()

    prefilled_equipment = None
    prefilled_building = None
    if initial_eq.isdigit():
        prefilled_equipment = (
            Equipment.objects.filter(pk=int(initial_eq), deleted=False)
            .select_related('equipment_type', 'building')
            .first()
        )
        if prefilled_equipment and not initial_b and prefilled_equipment.building_id:
            initial_b = str(prefilled_equipment.building_id)
    if initial_b.isdigit():
        prefilled_building = Building.objects.filter(pk=int(initial_b), deleted=False).first()

    obj = None
    if initial_eq or initial_b:
        class DummyObj:
            def __init__(self, eq_id, b_id):
                self.equipment_id = int(eq_id) if eq_id.isdigit() else None
                self.building_id = int(b_id) if b_id.isdigit() else None
                self.status = 'open'
                self.issue_type = 'breakdown'
                self.priority = 'medium'
                self.title = ''
                self.issue_description = ''
                self.due_date = None
                self.estimated_hours = None
        obj = DummyObj(initial_eq, initial_b)

    cancel_url = (
        reverse('backend:building_details', kwargs={'data_id': prefilled_building.id})
        if prefilled_building
        else reverse('backend:ticket_list')
    )

    context = {
        'action': 'Add',
        'is_edit': False,
        'equipments': equipments,
        'buildings': buildings,
        'obj': obj,
        'prefilled_equipment': prefilled_equipment,
        'prefilled_building': prefilled_building,
        'cancel_url': cancel_url,
        'issue_type_choices': Ticket.ISSUE_TYPE_CHOICES,
        'priority_choices': Ticket.PRIORITY_CHOICES,
        'status_choices': Ticket.STATUS_CHOICES,
        'existing_checklist': [],
    }
    return render(request, 'ticket/add.html', context)

@login_required
def ticket_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to update this ticket.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Ticket, pk=data_id, deleted=False)
    equipments = (
        Equipment.objects.filter(deleted=False, is_active=True)
        .select_related('equipment_type', 'building')
        .order_by('equipment_name', 'equipment_id')
    )
    buildings = Building.objects.filter(deleted=False, is_active=True).order_by('name')

    if request.method == 'POST':
        title = request.POST.get('title', '').strip()
        equipment_id = request.POST.get('equipment_id', '').strip()
        building_id = request.POST.get('building_id', '').strip()
        issue_type = request.POST.get('issue_type', 'breakdown').strip() or 'breakdown'
        priority = request.POST.get('priority', 'medium').strip() or 'medium'
        status = request.POST.get('status', obj.status).strip() or obj.status
        issue_description = request.POST.get('issue_description', '').strip()
        due_date_raw = request.POST.get('due_date', '').strip()
        est_hours_raw = request.POST.get('estimated_hours', '').strip()

        due_date_val = None
        if due_date_raw:
            try:
                parsed_d = datetime.strptime(due_date_raw, '%Y-%m-%d')
                due_date_val = timezone.make_aware(parsed_d) if timezone.is_naive(parsed_d) else parsed_d
            except ValueError:
                due_date_val = obj.due_date

        est_hours_val = None
        if est_hours_raw:
            try:
                est_hours_val = float(est_hours_raw)
            except ValueError:
                est_hours_val = obj.estimated_hours

        if not title or not equipment_id or not building_id:
            messages.error(request, 'Title, Equipment, and Building are required.')
        else:
            obj.title = title
            obj.equipment_id = equipment_id
            obj.building_id = building_id
            obj.issue_type = issue_type
            obj.priority = priority
            obj.status = status
            obj.issue_description = issue_description
            obj.due_date = due_date_val
            obj.estimated_hours = est_hours_val
            obj.updated_by = request.user
            obj.save()

            # Sync checklist items while preserving existing completion state for kept items
            submitted_ids = request.POST.getlist('checklist_item_ids')
            submitted_titles = request.POST.getlist('checklist_items')
            kept_ids = []
            for idx, raw_title in enumerate(submitted_titles, start=1):
                clean_title = (raw_title or '').strip()
                if not clean_title:
                    continue
                raw_id = submitted_ids[idx - 1].strip() if idx - 1 < len(submitted_ids) else ''
                if raw_id.isdigit():
                    existing_item = TicketTaskItem.objects.filter(pk=int(raw_id), ticket=obj, deleted=False).first()
                    if existing_item:
                        existing_item.title = clean_title[:255]
                        existing_item.order = idx
                        existing_item.updated_by = request.user
                        existing_item.save()
                        kept_ids.append(existing_item.id)
                        continue
                new_item = TicketTaskItem.objects.create(
                    ticket=obj,
                    title=clean_title[:255],
                    order=idx,
                    created_by=request.user,
                )
                kept_ids.append(new_item.id)
            TicketTaskItem.objects.filter(ticket=obj, deleted=False).exclude(id__in=kept_ids).update(
                deleted=True, is_active=False, updated_by=request.user
            )

            messages.success(request, 'Ticket updated successfully.')
            return redirect('backend:ticket_detail', data_id=obj.id)

    existing_checklist = list(obj.checklist_items.filter(deleted=False, is_active=True).order_by('order', 'id'))
    context = {
        'action': 'Update',
        'is_edit': True,
        'obj': obj,
        'equipments': equipments,
        'buildings': buildings,
        'prefilled_equipment': obj.equipment,
        'prefilled_building': obj.building,
        'cancel_url': reverse('backend:ticket_detail', kwargs={'data_id': obj.id}),
        'issue_type_choices': Ticket.ISSUE_TYPE_CHOICES,
        'priority_choices': Ticket.PRIORITY_CHOICES,
        'status_choices': Ticket.STATUS_CHOICES,
        'existing_checklist': existing_checklist,
    }
    return render(request, 'ticket/add.html', context)


@login_required
def ticket_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to change status of this ticket.')
        return render(request, '403.html', status=403)

    if request.method == 'POST':
        obj = get_object_or_404(Ticket, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Ticket status changed to {status_text} successfully.')
        return redirect('backend:ticket_list')

    messages.error(request, 'Invalid request method.')
    return redirect('backend:ticket_list')


@login_required
def maintanaince_create(request):
    if not checkUserPermission(request, "can_view", 'maintainance'):
        messages.error(request, 'You do not have permission to create maintainance task.')
        return render(request, '403.html', status=403)

    context = {}
    return render(request, 'maintainance/create.html', context) 


# ======================================== Maintenance Management ========================================

@login_required
def maintenance_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/maintenance/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)

    search = request.GET.get('search', '').strip()
    status = request.GET.get('status', '').strip()
    pretend_date_str = request.GET.get('pretend_date', '').strip()

    qs = Maintenance.objects.filter(deleted=False).select_related('building', 'created_by').prefetch_related(
        'equipment',
        'equipment__equipment_type',
        'equipment__building',
        'maintenance_records',
        'maintenance_records__equipment',
        'maintenance_records__equipment__equipment_type',
        'maintenance_records__equipment__building',
    )

    if search:
        qs = qs.filter(
            Q(maintenance_serial__icontains=search) |
            Q(qr_code__icontains=search) |
            Q(building__name__icontains=search) |
            Q(equipment__equipment_id__icontains=search) |
            Q(equipment__equipment_name__icontains=search) |
            Q(equipment__equipment_type__name__icontains=search) |
            Q(maintenance_records__equipment__equipment_id__icontains=search)
        ).distinct()

    if status == 'Active':
        qs = qs.filter(is_active=True)
    elif status == 'Inactive':
        qs = qs.filter(is_active=False)

    if request.GET.get('generate_due') == '1' and pretend_date_str:
        try:
            pretend_dt = datetime.strptime(pretend_date_str, '%Y-%m-%d').date()
            due_qs = qs.filter(is_active=True, equipment__next_service_due__lte=pretend_dt).distinct()
            due_count = due_qs.count()
            if due_count:
                messages.success(request, f'Found {due_count} due maintenance schedule(s) as of {pretend_dt.strftime("%d %b %Y")}.')
                qs = due_qs
            else:
                messages.info(request, f'No maintenance schedules are due as of {pretend_dt.strftime("%d %b %Y")}.')
        except ValueError:
            pass

    qs = qs.annotate(
        num_records=Count(
            'maintenance_records',
            filter=Q(maintenance_records__deleted=False),
            distinct=True
        )
    ).order_by('-created_at')

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    per_page = request.GET.get('per_page', 10)
    filtered_total_count = data_list.paginator.count

    # Attach primary equipment info for clean schedule table rendering
    for item in data_list:
        eqs = list(item.equipment.all())
        primary_eq = eqs[0] if eqs else None
        if not primary_eq:
            for rec in item.maintenance_records.all():
                if rec.equipment:
                    primary_eq = rec.equipment
                    break
        item.primary_eq = primary_eq
        item.extra_eq_count = max(0, len(eqs) - 1)

    today_iso = timezone.localdate().strftime('%Y-%m-%d')

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'per_page': per_page,
        'filtered_total_count': filtered_total_count,
        'search': search,
        'selected_status': status,
        'status_choices': ['Active', 'Inactive'],
        'pretend_date': pretend_date_str or today_iso,
    }
    return render(request, 'maintenance/list.html', context)


@login_required
def maintenance_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/maintenance/'):
        messages.error(request, 'You do not have permission to add a maintenance record.')
        return render(request, '403.html', status=403)

    buildings = Building.objects.filter(deleted=False, is_active=True).order_by('-created_at')
    technicians = Technician.objects.filter(deleted=False, is_active=True)
    
    logged_in_tech_id = None
    if request.user.is_authenticated:
        try:
            logged_in_tech_id = request.user.technician_profile.id
        except (AttributeError, Technician.DoesNotExist):
            pass

    if request.method == 'POST':
        building_id = request.POST.get('building')
        maintenance_serial = request.POST.get('maintenance_serial', '').strip() or None
        qr_code = request.POST.get('qr_code', '').strip() or None
        maintenance_date = request.POST.get('maintenance_date')

        # Gather dynamic records
        indexes = []
        for key in request.POST.keys():
            if key.startswith('records[') and key.endswith('][equipment_id]'):
                try:
                    idx = key.split('[')[1].split(']')[0]
                    indexes.append(int(idx))
                except (IndexError, ValueError):
                    pass
        indexes.sort()

        errors = []
        if not building_id:
            errors.append("Building is required.")
        else:
            try:
                building = Building.objects.get(id=building_id, deleted=False)
            except Building.DoesNotExist:
                errors.append("Selected building does not exist.")

        if not maintenance_date:
            errors.append("Maintenance Date is required.")

        if maintenance_serial and Maintenance.objects.filter(maintenance_serial=maintenance_serial, deleted=False).exists():
            errors.append("A maintenance record with this Serial already exists.")
        
        if qr_code and Maintenance.objects.filter(qr_code=qr_code, deleted=False).exists():
            errors.append("A maintenance record with this QR Code already exists.")

        submitted_records = []
        for idx in indexes:
            eq_id = request.POST.get(f'records[{idx}][equipment_id]')
            rec_type = request.POST.get(f'records[{idx}][record_type]', 'routine')
            maint_date = maintenance_date
            tech_id = request.POST.get(f'records[{idx}][technician_id]') or None
            work_desc = request.POST.get(f'records[{idx}][work_description]', '').strip()
            parts = request.POST.get(f'records[{idx}][parts_replaced]', '').strip()
            cost = request.POST.get(f'records[{idx}][cost]') or None
            duration = request.POST.get(f'records[{idx}][duration_hours]') or None
            status = request.POST.get(f'records[{idx}][maintainance_status]', 'pending')

            # Parse checklist components
            components_data = []
            comp_indexes = []
            prefix = f'records[{idx}][components]['
            for key in request.POST.keys():
                if key.startswith(prefix) and key.endswith('][name]'):
                    try:
                        c_idx = key.split(prefix)[1].split(']')[0]
                        comp_indexes.append(int(c_idx))
                    except (IndexError, ValueError):
                        pass
            comp_indexes.sort()

            for c_idx in comp_indexes:
                comp_name = request.POST.get(f'{prefix}{c_idx}][name]')
                is_checked = request.POST.get(f'{prefix}{c_idx}][is_checked]') in ('true', 'on')
                remark = request.POST.get(f'{prefix}{c_idx}][remark]', '').strip()
                suggestion = request.POST.get(f'{prefix}{c_idx}][suggestion]', '').strip()
                proceed = request.POST.get(f'{prefix}{c_idx}][proceed]') == 'true'
                components_data.append({
                    'name': comp_name,
                    'is_checked': is_checked,
                    'remark': remark,
                    'suggestion': suggestion,
                    'proceed': proceed
                })

            record_data = {
                'equipment_id': eq_id,
                'record_type': rec_type,
                'maintenance_date': maint_date,
                'technician_id': tech_id,
                'work_description': work_desc,
                'parts_replaced': parts,
                'cost': cost,
                'duration_hours': duration,
                'maintainance_status': status,
                'components': components_data,
            }
            submitted_records.append(record_data)

            if not eq_id:
                errors.append(f"Row {idx + 1}: Equipment is required.")
            else:
                if not Equipment.objects.filter(id=eq_id, building_id=building_id, deleted=False).exists():
                    errors.append(f"Row {idx + 1}: Selected equipment does not belong to the selected building.")
            if not maint_date:
                errors.append(f"Row {idx + 1}: Maintenance Date is required.")
            if not work_desc:
                errors.append(f"Row {idx + 1}: Work Description / Issue is required.")

        if not indexes or not submitted_records:
            errors.append("Please select an equipment and enter the maintenance details.")

        if errors:
            for error in errors:
                messages.error(request, error)
            
            # Pass back context to reconstruct the page
            context = {
                'action': 'Add',
                'buildings': buildings,
                'technicians': technicians,
                'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
                'status_choices': MaintenanceRecord.MAINTAIANCE_STATUS_CHOICES,
                'submitted_building_id': building_id,
                'submitted_serial': maintenance_serial,
                'submitted_qr_code': qr_code,
                'submitted_date': maintenance_date,
                'default_date': date.today().strftime('%Y-%m-%d'),
                'submitted_records': json.dumps(submitted_records),
                'logged_in_tech_id': logged_in_tech_id,
            }
            return render(request, 'maintenance/add.html', context)

        from django.db import transaction
        try:
            with transaction.atomic():
                maintenance = Maintenance(
                    building_id=building_id,
                    maintenance_serial=maintenance_serial,
                    qr_code=qr_code,
                    created_by=request.user,
                    is_active=True
                )
                maintenance.save()

                for rec in submitted_records:
                    tech_user = None
                    assigned_tech = None
                    if rec['technician_id']:
                        try:
                            assigned_tech = Technician.objects.get(id=rec['technician_id'])
                            tech_user = assigned_tech.user
                        except Technician.DoesNotExist:
                            pass
                    record = MaintenanceRecord(
                        equipment_id=rec['equipment_id'],
                        maintenance=maintenance,
                        record_type=rec['record_type'],
                        maintenance_date=rec['maintenance_date'],
                        technician=tech_user,
                        assigned_to=assigned_tech,
                        work_description=rec['work_description'],
                        parts_replaced=rec['parts_replaced'],
                        cost=rec['cost'],
                        duration_hours=rec['duration_hours'],
                        maintainance_status=rec['maintainance_status'],
                        created_by=request.user
                    )
                    record.save()
                    maintenance.equipment.add(rec['equipment_id'])

                    # Save components checklist
                    for comp in rec.get('components', []):
                        MaintenanceComponent.objects.create(
                            maintenance_record=record,
                            name=comp['name'],
                            is_checked=comp['is_checked'],
                            remark=comp['remark'],
                            suggestion=comp['suggestion'],
                            proceed=comp.get('proceed', False),
                            created_by=request.user
                        )
                        # Ensure exists in EquipmentComponents
                        if not EquipmentComponents.objects.filter(equipment_id=rec['equipment_id'], name__iexact=comp['name'], deleted=False).exists():
                            EquipmentComponents.objects.create(
                                equipment_id=rec['equipment_id'],
                                name=comp['name'],
                                created_by=request.user
                            )

                messages.success(request, 'Maintenance session created successfully.')
                return redirect('backend:maintenance_detail', data_id=maintenance.id)
        except Exception as e:
            messages.error(request, f"Error saving maintenance session: {str(e)}")

    submitted_building_id = request.GET.get('building')
    submitted_records = []
    if request.GET.get('equipment'):
        record = {'equipment_id': request.GET.get('equipment')}
        if logged_in_tech_id:
            record['technician_id'] = str(logged_in_tech_id)
        submitted_records.append(record)

    context = {
        'action': 'Add',
        'buildings': buildings,
        'technicians': technicians,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'status_choices': MaintenanceRecord.MAINTAIANCE_STATUS_CHOICES,
        'default_date': date.today().strftime('%Y-%m-%d'),
        'submitted_building_id': submitted_building_id,
        'submitted_records': json.dumps(submitted_records),
        'logged_in_tech_id': logged_in_tech_id,
    }
    return render(request, 'maintenance/add.html', context)


@login_required
def maintenance_detail(request, data_id):
    obj = get_object_or_404(Maintenance, pk=data_id, deleted=False)

    # Get user's technician profile
    technician = None
    try:
        technician = request.user.technician_profile
    except Exception:
        pass

    # Check if technician is assigned to any records in this maintenance session
    is_assigned = False
    if technician:
        is_assigned = obj.maintenance_records.filter(assigned_to=technician, deleted=False).exists()

    has_general_permission = checkUserPermission(request, 'can_view', '/backend/maintenance/')

    if not has_general_permission and not is_assigned:
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)

    session_records = obj.maintenance_records.filter(deleted=False).select_related(
        'equipment', 'equipment__building', 'equipment__equipment_type',
        'technician', 'assigned_to', 'ticket'
    ).order_by('-maintenance_date', '-created_at')

    session_date = None
    first_session_rec = session_records.first()
    if first_session_rec:
        session_date = first_session_rec.maintenance_date

    # Since there will only be one equipment for one maintenance session,
    # find the equipment of this session.
    equipment = None
    if first_session_rec:
        equipment = first_session_rec.equipment
    else:
        equipment = obj.equipment.first()

    # Query all maintenance records for that equipment (all times history)
    if equipment:
        records = MaintenanceRecord.objects.filter(equipment=equipment, deleted=False).select_related(
            'equipment', 'equipment__building', 'equipment__equipment_type',
            'technician', 'assigned_to', 'ticket'
        ).order_by('-maintenance_date', '-created_at')
    else:
        records = session_records

    # If technician and no general permission, only show records assigned to them
    if technician and not has_general_permission:
        records = records.filter(assigned_to=technician)

    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'delete_record':
            if not checkUserPermission(request, 'can_update', '/backend/maintenance/'):
                messages.error(request, 'You do not have permission to delete this record.')
                return redirect('backend:maintenance_detail', data_id=obj.id)
            record_id = request.POST.get('record_id')
            if record_id:
                record = get_object_or_404(MaintenanceRecord, pk=record_id, maintenance=obj)
                record.deleted = True
                record.updated_by = request.user
                record.save()
                messages.success(request, 'Maintenance record deleted successfully.')
            return redirect('backend:maintenance_detail', data_id=obj.id)

    qr_base64 = None
    if obj.qr_code:
        try:
            scan_url = request.build_absolute_uri(reverse('backend:maintanaince_scan', kwargs={'qr_code': obj.qr_code}))
            img = generate_qr_code(scan_url)
            buffer = BytesIO()
            img.save(buffer, format="PNG")
            qr_base64 = base64.b64encode(buffer.getvalue()).decode()
        except Exception as e:
            pass

    context = {
        'obj': obj,
        'records': records,
        'session_date': session_date,
        'qr_base64': qr_base64,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'status_choices': MaintenanceRecord.MAINTAIANCE_STATUS_CHOICES,
    }
    return render(request, 'maintenance/detail.html', context)


@login_required
def maintenance_record_detail(request, record_id):
    record = get_object_or_404(
        MaintenanceRecord.objects.select_related(
            'equipment', 'equipment__building', 'equipment__equipment_type',
            'maintenance', 'maintenance__building', 'ticket',
            'technician', 'assigned_to', 'reviewed_by', 'created_by', 'updated_by'
        ),
        pk=record_id,
        deleted=False
    )

    technician = None
    try:
        technician = request.user.technician_profile
    except Exception:
        pass

    is_assigned = (technician is not None and record.assigned_to == technician)
    has_general_permission = checkUserPermission(request, 'can_view', '/backend/maintenance/')

    if not has_general_permission and not is_assigned:
        messages.error(request, 'You do not have permission to view this maintenance record.')
        return render(request, '403.html', status=403)

    components = record.components.filter(deleted=False)
    can_review = checkUserPermission(request, 'can_edit', '/backend/maintenance/') or request.user.is_superuser

    context = {
        'record': record,
        'obj': record,
        'components': components,
        'can_review': can_review,
    }
    return render(request, 'maintenance/record_detail.html', context)


@login_required
def maintenance_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/maintenance/'):
        messages.error(request, 'You do not have permission to edit this record.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Maintenance, pk=data_id, deleted=False)
    buildings = Building.objects.filter(deleted=False, is_active=True).order_by('-created_at')
    technicians = Technician.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        building_id = request.POST.get('building')
        maintenance_serial = request.POST.get('maintenance_serial', '').strip() or None
        qr_code = request.POST.get('qr_code', '').strip() or None
        maintenance_date = request.POST.get('maintenance_date')

        # Gather dynamic records
        indexes = []
        for key in request.POST.keys():
            if key.startswith('records[') and key.endswith('][equipment_id]'):
                try:
                    idx = key.split('[')[1].split(']')[0]
                    indexes.append(int(idx))
                except (IndexError, ValueError):
                    pass
        indexes.sort()

        errors = []
        if not building_id:
            errors.append("Building is required.")
        else:
            try:
                building = Building.objects.get(id=building_id, deleted=False)
            except Building.DoesNotExist:
                errors.append("Selected building does not exist.")

        if not maintenance_date:
            errors.append("Maintenance Date is required.")

        if maintenance_serial and Maintenance.objects.filter(maintenance_serial=maintenance_serial, deleted=False).exclude(pk=data_id).exists():
            errors.append("A maintenance record with this Serial already exists.")
        
        if qr_code and Maintenance.objects.filter(qr_code=qr_code, deleted=False).exclude(pk=data_id).exists():
            errors.append("A maintenance record with this QR Code already exists.")

        submitted_records = []
        for idx in indexes:
            rec_id = request.POST.get(f'records[{idx}][id]') or None
            eq_id = request.POST.get(f'records[{idx}][equipment_id]')
            rec_type = request.POST.get(f'records[{idx}][record_type]', 'routine')
            maint_date = maintenance_date
            tech_id = request.POST.get(f'records[{idx}][technician_id]') or None
            work_desc = request.POST.get(f'records[{idx}][work_description]', '').strip()
            parts = request.POST.get(f'records[{idx}][parts_replaced]', '').strip()
            cost = request.POST.get(f'records[{idx}][cost]') or None
            duration = request.POST.get(f'records[{idx}][duration_hours]') or None
            status = request.POST.get(f'records[{idx}][maintainance_status]', 'pending')

            # Parse checklist components
            components_data = []
            comp_indexes = []
            prefix = f'records[{idx}][components]['
            for key in request.POST.keys():
                if key.startswith(prefix) and key.endswith('][name]'):
                    try:
                        c_idx = key.split(prefix)[1].split(']')[0]
                        comp_indexes.append(int(c_idx))
                    except (IndexError, ValueError):
                        pass
            comp_indexes.sort()

            for c_idx in comp_indexes:
                comp_name = request.POST.get(f'{prefix}{c_idx}][name]')
                is_checked = request.POST.get(f'{prefix}{c_idx}][is_checked]') in ('true', 'on')
                remark = request.POST.get(f'{prefix}{c_idx}][remark]', '').strip()
                suggestion = request.POST.get(f'{prefix}{c_idx}][suggestion]', '').strip()
                proceed = request.POST.get(f'{prefix}{c_idx}][proceed]') == 'true'
                components_data.append({
                    'name': comp_name,
                    'is_checked': is_checked,
                    'remark': remark,
                    'suggestion': suggestion,
                    'proceed': proceed
                })

            record_data = {
                'id': rec_id,
                'equipment_id': eq_id,
                'record_type': rec_type,
                'maintenance_date': maint_date,
                'technician_id': tech_id,
                'work_description': work_desc,
                'parts_replaced': parts,
                'cost': cost,
                'duration_hours': duration,
                'maintainance_status': status,
                'components': components_data,
            }
            submitted_records.append(record_data)

            if not eq_id:
                errors.append(f"Row {idx + 1}: Equipment is required.")
            else:
                if not Equipment.objects.filter(id=eq_id, building_id=building_id, deleted=False).exists():
                    errors.append(f"Row {idx + 1}: Selected equipment does not belong to the selected building.")
            if not maint_date:
                errors.append(f"Row {idx + 1}: Maintenance Date is required.")
            if not work_desc:
                errors.append(f"Row {idx + 1}: Work Description / Issue is required.")

        if not indexes or not submitted_records:
            errors.append("Please select an equipment and enter the maintenance details.")

        if errors:
            for error in errors:
                messages.error(request, error)
            
            # Pass back context to reconstruct
            context = {
                'action': 'Update',
                'obj': obj,
                'buildings': buildings,
                'technicians': technicians,
                'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
                'status_choices': MaintenanceRecord.MAINTAIANCE_STATUS_CHOICES,
                'submitted_building_id': building_id,
                'submitted_serial': maintenance_serial,
                'submitted_qr_code': qr_code,
                'submitted_date': maintenance_date,
                'submitted_records': json.dumps(submitted_records),
            }
            return render(request, 'maintenance/update.html', context)

        from django.db import transaction
        try:
            with transaction.atomic():
                obj.building_id = building_id
                obj.maintenance_serial = maintenance_serial or obj.maintenance_serial or f"MT-{uuid.uuid4().hex[:8].upper()}"
                obj.qr_code = qr_code or obj.qr_code or f"QR-MT-{uuid.uuid4().hex[:8].upper()}"
                obj.updated_by = request.user
                obj.save()

                # Get existing record IDs in database
                existing_db_records = obj.maintenance_records.filter(deleted=False)
                existing_db_ids = set(str(r.id) for r in existing_db_records)

                submitted_ids = set(rec['id'] for rec in submitted_records if rec['id'])

                # Delete those that were removed
                deleted_ids = existing_db_ids - submitted_ids
                if deleted_ids:
                    MaintenanceRecord.objects.filter(id__in=deleted_ids).update(deleted=True, updated_by=request.user)

                obj.equipment.clear()

                for rec in submitted_records:
                    tech_user = None
                    assigned_tech = None
                    if rec['technician_id']:
                        try:
                            assigned_tech = Technician.objects.get(id=rec['technician_id'])
                            tech_user = assigned_tech.user
                        except Technician.DoesNotExist:
                            pass

                    if rec['id'] and rec['id'] in existing_db_ids:
                        # Update existing
                        record = MaintenanceRecord.objects.get(id=rec['id'])
                        record.equipment_id = rec['equipment_id']
                        record.record_type = rec['record_type']
                        record.maintenance_date = rec['maintenance_date']
                        record.technician = tech_user
                        record.assigned_to = assigned_tech
                        record.work_description = rec['work_description']
                        record.parts_replaced = rec['parts_replaced']
                        record.cost = rec['cost']
                        record.duration_hours = rec['duration_hours']
                        record.maintainance_status = rec['maintainance_status']
                        record.updated_by = request.user
                        record.save()
                    else:
                        # Create new
                        record = MaintenanceRecord(
                            equipment_id=rec['equipment_id'],
                            maintenance=obj,
                            record_type=rec['record_type'],
                            maintenance_date=rec['maintenance_date'],
                            technician=tech_user,
                            assigned_to=assigned_tech,
                            work_description=rec['work_description'],
                            parts_replaced=rec['parts_replaced'],
                            cost=rec['cost'],
                            duration_hours=rec['duration_hours'],
                            maintainance_status=rec['maintainance_status'],
                            created_by=request.user
                        )
                        record.save()

                    # Save components checklist: delete old components first, then recreate
                    record.components.all().delete()
                    for comp in rec.get('components', []):
                        MaintenanceComponent.objects.create(
                            maintenance_record=record,
                            name=comp['name'],
                            is_checked=comp['is_checked'],
                            remark=comp['remark'],
                            suggestion=comp['suggestion'],
                            proceed=comp.get('proceed', False),
                            created_by=request.user
                        )
                        # Ensure exists in EquipmentComponents
                        if not EquipmentComponents.objects.filter(equipment_id=rec['equipment_id'], name__iexact=comp['name'], deleted=False).exists():
                            EquipmentComponents.objects.create(
                                equipment_id=rec['equipment_id'],
                                name=comp['name'],
                                created_by=request.user
                            )

                    obj.equipment.add(rec['equipment_id'])

                messages.success(request, 'Maintenance session updated successfully.')
                return redirect('backend:maintenance_detail', data_id=obj.id)
        except Exception as e:
            messages.error(request, f"Error updating maintenance session: {str(e)}")

    # GET request: serialize current records
    existing_records = []
    submitted_date = date.today().strftime('%Y-%m-%d')
    for r in obj.maintenance_records.filter(deleted=False):
        if r.maintenance_date:
            submitted_date = r.maintenance_date.strftime('%Y-%m-%d')
            
        # Serialize existing components
        components_data = []
        for comp in r.components.filter(deleted=False):
            components_data.append({
                'name': comp.name,
                'is_checked': comp.is_checked,
                'remark': comp.remark or '',
                'suggestion': comp.suggestion or '',
                'proceed': comp.proceed,
            })
            
        existing_records.append({
            'id': str(r.id),
            'equipment_id': r.equipment_id,
            'record_type': r.record_type,
            'maintenance_date': r.maintenance_date.strftime('%Y-%m-%d') if r.maintenance_date else '',
            'technician_id': r.assigned_to_id or '',
            'work_description': r.work_description,
            'parts_replaced': r.parts_replaced,
            'cost': str(r.cost) if r.cost is not None else '',
            'duration_hours': str(r.duration_hours) if r.duration_hours is not None else '',
            'maintainance_status': r.maintainance_status,
            'components': components_data,
        })

    context = {
        'action': 'Update',
        'obj': obj,
        'buildings': buildings,
        'technicians': technicians,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'status_choices': MaintenanceRecord.MAINTAIANCE_STATUS_CHOICES,
        'submitted_building_id': str(obj.building_id) if obj.building_id else '',
        'submitted_serial': obj.maintenance_serial,
        'submitted_qr_code': obj.qr_code,
        'submitted_date': submitted_date,
        'submitted_records': json.dumps(existing_records),
    }
    return render(request, 'maintenance/update.html', context)


@login_required
def maintenance_status(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/maintenance/'):
        messages.error(request, 'You do not have permission to change status of this maintenance record.')
        return render(request, '403.html', status=403)

    if request.method == 'POST':
        obj = get_object_or_404(Maintenance, pk=data_id, deleted=False)
        obj.is_active = not obj.is_active
        obj.updated_by = request.user
        obj.save()
        status_text = "Active" if obj.is_active else "Inactive"
        messages.success(request, f'Maintenance status changed to {status_text} successfully.')
        return redirect('backend:maintenance')

    messages.error(request, 'Invalid request method.')
    return redirect('backend:maintenance')


@login_required
@require_GET
def get_building_equipment(request):
    """AJAX endpoint to get equipment for a selected building"""
    building_id = request.GET.get('building_id')
    
    if not building_id:
        return JsonResponse({'equipment': []})
    
    try:
        # Fetch equipment for the building
        equipment_list = Equipment.objects.filter(
            building_id=building_id,
            deleted=False,
            is_active=True
        ).select_related('equipment_type').order_by('equipment_id')
        
        equipment = [
            {
                'id': item.id,
                'equipment_id': item.equipment_id,
                'display_name': f"{item.equipment_id} - {item.equipment_type.name}" if (item.equipment_type and item.equipment_type.name) else item.equipment_id,
            }
            for item in equipment_list
        ]
        
        return JsonResponse({'equipment': equipment})
    except Exception as e:
        return JsonResponse({'error': str(e), 'equipment': []}, status=400)


@login_required
def maintanaince_scan(request, qr_code=None):
    if not qr_code:
        qr_code = request.GET.get('qr_code', '').strip()
    
    if not qr_code:
        messages.error(request, "No QR Code provided.")
        return redirect('backend:dash_board')
        
    maintenance = get_object_or_404(Maintenance, qr_code=qr_code, deleted=False)
    
    # Check if user has a technician profile
    technician = None
    try:
        technician = request.user.technician_profile
    except Exception:
        pass

    if technician:
        # Check if the technician is assigned to any records in this maintenance session
        assigned_records_exist = maintenance.maintenance_records.filter(
            assigned_to=technician, deleted=False
        ).exists()
        
        if assigned_records_exist:
            # Redirect to the maintenance detail page where they will see their assigned tasks
            return redirect('backend:maintenance_detail', data_id=maintenance.id)
        else:
            # If they are not assigned to this session, check if they have general view permission
            if checkUserPermission(request, 'can_view', '/backend/maintenance/'):
                return redirect('backend:maintenance_detail', data_id=maintenance.id)
            else:
                messages.error(request, "You are not assigned to this maintenance session.")
                return redirect('backend:dash_board')
    else:
        # If they are an administrator or other user, check if they have general view permission
        if checkUserPermission(request, 'can_view', '/backend/maintenance/'):
            return redirect('backend:maintenance_detail', data_id=maintenance.id)
        else:
            messages.error(request, "You do not have permission to view this maintenance session.")
            return redirect('backend:dash_board') 


@login_required
@require_GET
def get_equipment_components(request):
    """AJAX endpoint to get components for a selected equipment"""
    equipment_id = request.GET.get('equipment_id')
    if not equipment_id:
        return JsonResponse({'components': []})
    try:
        components_list = EquipmentComponents.objects.filter(
            equipment_id=equipment_id,
            deleted=False,
            is_active=True
        ).values('id', 'name', 'description').order_by('name')
        
        components = [
            {
                'id': item['id'],
                'name': item['name'],
                'description': item['description'],
            }
            for item in components_list
        ]
        return JsonResponse({'components': components})
    except Exception as e:
        return JsonResponse({'error': str(e), 'components': []}, status=400)


def _start_ticket_work(ticket_obj, accepted_act, actor=None, verified_via_qr=False):
    """Transition an accepted ticket assignment into In Progress and record TicketActivityLog."""
    now = timezone.now()
    accepted_act.completion_status = 'in_progress'
    if not accepted_act.started_at:
        accepted_act.started_at = now
    if actor and getattr(actor, 'is_authenticated', False):
        accepted_act.updated_by = actor
    accepted_act.save()

    ticket_obj.status = 'in_progress'
    if actor and getattr(actor, 'is_authenticated', False):
        ticket_obj.updated_by = actor
    ticket_obj.save()

    if ticket_obj.equipment and ticket_obj.equipment.status == 'active':
        ticket_obj.equipment.status = 'under_maintenance'
        ticket_obj.equipment.save()

    tech_obj = accepted_act.technician
    eq_label = (
        ticket_obj.equipment.equipment_name or ticket_obj.equipment.equipment_id
        if ticket_obj.equipment else 'equipment'
    )
    if verified_via_qr:
        desc = f"{tech_obj.display_name} verified equipment QR & PIN and started maintenance work on {eq_label}."
    else:
        desc = f"{tech_obj.display_name} started maintenance work on {eq_label}."

    TicketActivityLog.objects.create(
        ticket=ticket_obj,
        actor=actor if (actor and getattr(actor, 'is_authenticated', False)) else None,
        action_type='work_started',
        description=desc,
        metadata={
            'technician_id': tech_obj.id,
            'technician_name': tech_obj.display_name,
            'verified_via_qr': verified_via_qr,
        },
    )


def static_otp_verification(request):
    building_id = (request.POST.get('building_id') or request.GET.get('building') or '').strip()
    equipment_id = (request.POST.get('equipment_id') or request.GET.get('equipment') or '').strip()
    ticket_id = (request.POST.get('ticket_id') or request.GET.get('ticket') or '').strip()
    technician_id = (request.POST.get('technician_id') or request.GET.get('technician') or '').strip()

    ticket_obj = None
    technician_obj = None
    equipment_obj = None
    if ticket_id.isdigit():
        ticket_obj = Ticket.objects.filter(pk=int(ticket_id), deleted=False).select_related('equipment', 'building').first()
    if technician_id.isdigit():
        technician_obj = Technician.objects.filter(pk=int(technician_id), deleted=False).first()
    if equipment_id.isdigit():
        equipment_obj = Equipment.objects.filter(pk=int(equipment_id), deleted=False).select_related('building', 'equipment_type').first()
    elif ticket_obj and ticket_obj.equipment:
        equipment_obj = ticket_obj.equipment
        equipment_id = str(equipment_obj.id)
    if not building_id and equipment_obj and equipment_obj.building_id:
        building_id = str(equipment_obj.building_id)

    if request.method == 'POST' and ticket_obj:
        pin = (request.POST.get('pin') or '').strip()
        if pin != '1234':
            messages.error(request, 'Incorrect 4-digit PIN. Please try again.')
        else:
            accepted_qs = AssignActivities.objects.filter(
                ticket=ticket_obj, is_active=True, status='accepted'
            ).select_related('technician')
            if technician_obj:
                accepted_act = accepted_qs.filter(technician=technician_obj).first()
            else:
                accepted_act = accepted_qs.first()

            if not accepted_act:
                messages.error(request, 'Technician must accept this task first before starting work.')
                if technician_obj:
                    return redirect(f"{reverse('backend:technician_details', kwargs={'data_id': technician_obj.id})}?tab=tasks")
                return redirect('backend:ticket_detail', data_id=ticket_obj.id)

            if ticket_obj.status != 'in_progress' or not accepted_act.started_at:
                _start_ticket_work(ticket_obj, accepted_act, actor=request.user, verified_via_qr=True)
            target_tech_id = technician_obj.id if technician_obj else accepted_act.technician_id
            request.session[f'ticket_pin_verified_{ticket_obj.id}_{target_tech_id}'] = True
            messages.success(
                request,
                f'PIN verified! You can now complete the checklist and submit Ticket #{ticket_obj.id}.'
            )
            return redirect('backend:technician_task_detail', tech_id=target_tech_id, ticket_id=ticket_obj.id)

    context = {
        'building_id': building_id,
        'equipment_id': equipment_id,
        'ticket_id': ticket_id,
        'technician_id': technician_id,
        'ticket_obj': ticket_obj,
        'technician_obj': technician_obj,
        'equipment_obj': equipment_obj,
    }
    return render(request, 'maintenance/static_otp.html', context)


@require_GET
def get_division_subdivisions(request):
    """AJAX endpoint to get subdivisions for a selected division"""
    division_id = request.GET.get('division_id')
    if not division_id:
        return JsonResponse({'subdivisions': []})
    try:
        subdivision_list = SubDivision.objects.filter(
            division_id=division_id,
            deleted=False,
            is_active=True
        ).values('id', 'name').order_by('name')
        subdivisions = [
            {'id': item['id'], 'name': item['name']}
            for item in subdivision_list
        ]
        return JsonResponse({'subdivisions': subdivisions})
    except Exception as e:
        return JsonResponse({'error': str(e), 'subdivisions': []}, status=400)


@require_GET
def get_subdivision_sections(request):
    """AJAX endpoint to get sections for a selected subdivision"""
    subdivision_id = request.GET.get('subdivision_id')
    if not subdivision_id:
        return JsonResponse({'sections': []})
    try:
        section_list = Section.objects.filter(
            subdivision_id=subdivision_id,
            deleted=False,
            is_active=True
        ).values('id', 'name').order_by('name')
        sections = [
            {'id': item['id'], 'name': item['name']}
            for item in section_list
        ]
        return JsonResponse({'sections': sections})
    except Exception as e:
        return JsonResponse({'error': str(e), 'sections': []}, status=400)


@login_required
@require_POST
def quick_create_division(request):
    """AJAX endpoint to quick-create a division from the searchable dropdown"""
    name = request.POST.get('name', '').strip()
    if not name:
        return JsonResponse({'success': False, 'error': 'Division name is required.'}, status=400)

    # Check if division with this name already exists (case-insensitive)
    existing = Division.objects.filter(name__iexact=name).first()
    if existing:
        if existing.deleted:
            existing.deleted = False
            existing.is_active = True
            existing.updated_by = request.user
            existing.save()
        return JsonResponse({
            'success': True,
            'id': existing.id,
            'name': existing.name,
            'code': existing.code,
            'already_existed': True
        })

    words = [w for w in name.split() if w]
    if len(words) >= 2:
        abbr = ''.join([w[0].upper() for w in words])[:6]
    else:
        abbr = name[:4].upper()

    base_code = f"DIV-{abbr}"
    code = base_code
    counter = 1
    while Division.objects.filter(code=code).exists():
        code = f"{base_code}-{counter}"
        counter += 1

    division = Division.objects.create(
        name=name,
        code=code,
        is_active=True,
        deleted=False,
        created_by=request.user
    )

    return JsonResponse({
        'success': True,
        'id': division.id,
        'name': division.name,
        'code': division.code
    })


@login_required
@require_POST
def quick_create_subdivision(request):
    """AJAX endpoint to quick-create a subdivision under a division"""
    division_id = request.POST.get('division_id', '').strip()
    name = request.POST.get('name', '').strip()

    if not division_id:
        return JsonResponse({'success': False, 'error': 'Parent division is required.'}, status=400)
    if not name:
        return JsonResponse({'success': False, 'error': 'Subdivision name is required.'}, status=400)

    division = Division.objects.filter(id=division_id, deleted=False).first()
    if not division:
        return JsonResponse({'success': False, 'error': 'Selected division not found.'}, status=404)

    # Check if a subdivision with this name already exists under this division (case-insensitive)
    existing = SubDivision.objects.filter(division=division, name__iexact=name).first()
    if existing:
        if existing.deleted:
            existing.deleted = False
            existing.is_active = True
            existing.updated_by = request.user
            existing.save()
        return JsonResponse({
            'success': True,
            'id': existing.id,
            'name': existing.name,
            'code': existing.code,
            'already_existed': True
        })

    words = [w for w in name.split() if w]
    if len(words) >= 2:
        abbr = ''.join([w[0].upper() for w in words])[:6]
    else:
        abbr = name[:4].upper()

    base_code = f"SUB-{abbr}"
    code = base_code
    counter = 1
    while SubDivision.objects.filter(code=code).exists():
        code = f"{base_code}-{counter}"
        counter += 1

    subdivision = SubDivision.objects.create(
        division=division,
        name=name,
        code=code,
        is_active=True,
        deleted=False,
        created_by=request.user
    )

    return JsonResponse({
        'success': True,
        'id': subdivision.id,
        'name': subdivision.name,
        'code': subdivision.code
    })


@login_required
@require_POST
def quick_create_section(request):
    """AJAX endpoint to quick-create a section under a subdivision"""
    subdivision_id = request.POST.get('subdivision_id', '').strip()
    name = request.POST.get('name', '').strip()

    if not subdivision_id:
        return JsonResponse({'success': False, 'error': 'Parent subdivision is required.'}, status=400)
    if not name:
        return JsonResponse({'success': False, 'error': 'Section name is required.'}, status=400)

    subdivision = SubDivision.objects.filter(id=subdivision_id, deleted=False).first()
    if not subdivision:
        return JsonResponse({'success': False, 'error': 'Selected subdivision not found.'}, status=404)

    # Check if a section with this name already exists under this subdivision (case-insensitive)
    existing = Section.objects.filter(subdivision=subdivision, name__iexact=name).first()
    if existing:
        if existing.deleted:
            existing.deleted = False
            existing.is_active = True
            existing.updated_by = request.user
            existing.save()
        return JsonResponse({
            'success': True,
            'id': existing.id,
            'name': existing.name,
            'code': existing.code,
            'already_existed': True
        })

    words = [w for w in name.split() if w]
    if len(words) >= 2:
        abbr = ''.join([w[0].upper() for w in words])[:6]
    else:
        abbr = name[:4].upper()

    base_code = f"SEC-{abbr}"
    code = base_code
    counter = 1
    while Section.objects.filter(code=code).exists():
        code = f"{base_code}-{counter}"
        counter += 1

    section = Section.objects.create(
        subdivision=subdivision,
        name=name,
        code=code,
        is_active=True,
        deleted=False,
        created_by=request.user
    )

    return JsonResponse({
        'success': True,
        'id': section.id,
        'name': section.name,
        'code': section.code
    })





