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
from django.views.decorators.http import require_GET
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
    Ticket, TicketComment, TicketActivityLog, AssignTechnician, Technician, AssignActivities,
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
    items_per_page, max_pages = 10, 10
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
    qs = Division.objects.filter(deleted=False)
    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(code__icontains=search))

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'search': search,
    }
    return render(request, 'division/list.html', context)


@login_required
def division_add(request):
    if not checkUserPermission(request, "can_add", "/backend/location/division/"):
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        code = request.POST.get('code', '').strip()
        location = request.POST.get('location', '').strip()

        if not name or not code:
            messages.error(request, 'Name and Code are required.')
        elif Division.objects.filter(code=code, deleted=False).exists():
            messages.error(request, 'A division with this code already exists.')
        else:
            Division.objects.create(
                name=name, code=code,
                is_active=True, 
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
        location = request.POST.get('location', '').strip()

        if not name or not code:
            messages.error(request, 'Name and Code are required.')
        elif Division.objects.filter(code=code, deleted=False).exclude(pk=data_id).exists():
            messages.error(request, 'A division with this code already exists.')
        else:
            obj.name = name
            obj.code = code
            obj.is_active = True
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
    qs = SubDivision.objects.filter(deleted=False).select_related('division')

    if search:
        qs = qs.filter(Q(name__icontains=search))
    if division_id:
        qs = qs.filter(division_id=division_id)

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'search': search,
        'division_id': division_id,
        'divisions': Division.objects.filter(deleted=False, is_active=True),
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
        location = request.POST.get('location', '').strip()

        if not division_id or not name or not code:
            messages.error(request, 'Division, Name and Code are required.')
        else:
            SubDivision.objects.create(
                division_id=division_id, 
                name=name,
                code=code,
                is_active=True, 
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
        location = request.POST.get('location', '').strip()

        if not division_id or not name or not code:
            messages.error(request, 'Division, Name and Code are required.')
        else:
            obj.division_id = division_id
            obj.name = name
            obj.code = code
            obj.is_active = True
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
    qs = Section.objects.filter(deleted=False).select_related('subdivision')
    if search:
        qs = qs.filter(Q(name__icontains=search))
    if subdivision_id:
        qs = qs.filter(subdivision_id=subdivision_id)

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'search': search,
        'subdivision_id': subdivision_id,
        'subdivisions': SubDivision.objects.filter(deleted=False, is_active=True),
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
        location = request.POST.get('location', '').strip()

        if not subdivision_id or not name or not code:
            messages.error(request, 'Sub-Division, Name and Code are required.')
        elif Section.objects.filter(subdivision_id=subdivision_id, code=code).exists():
            messages.error(request, 'A section with this code already exists in this sub-division.')
        else:
            Section.objects.create(
                subdivision_id=subdivision_id, name=name, code=code,
                is_active=True, created_by=request.user,
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
        location = request.POST.get('location', '').strip()

        if not subdivision_id or not name or not code:
            messages.error(request, 'Sub-Division, Name and Code are required.')
        elif Section.objects.filter(subdivision_id=subdivision_id, code=code).exclude(pk=data_id).exists():
            messages.error(request, 'A section with this code already exists in this sub-division.')
        else:
            obj.subdivision_id = subdivision_id
            obj.name = name
            obj.code = code
            obj.is_active = True
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

    qs = Building.objects.filter(deleted=False).select_related('division', 'subdivision', 'section')
    
    if search:
        qs = qs.filter(Q(name__icontains=search) | Q(building_id__icontains=search) | Q(surname__icontains=search))
    if division_id:
        qs = qs.filter(division_id=division_id)
    if subdivision_id:
        qs = qs.filter(subdivision_id=subdivision_id)
    if section_id:
        qs = qs.filter(section_id=section_id)
    if status:
        qs = qs.filter(status=status)

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)

    # Fetch filter options
    divisions = Division.objects.filter(deleted=False, is_active=True)
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True)
    sections = Section.objects.filter(deleted=False, is_active=True)
    status_choices = [choice[0] for choice in Building.STATUS_CHOICES]

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
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

    divisions = Division.objects.filter(deleted=False, is_active=True)
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True)
    sections = Section.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        building_id = request.POST.get('building_id', '').strip()
        name = request.POST.get('name', '').strip()
        surname = request.POST.get('surname', '').strip()
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

        if not building_id:
            messages.error(request, 'Building ID is required.')
        elif not name:
            messages.error(request, 'Building Name is required.')
        elif not section_id:
            messages.error(request, 'Section is required.')
        elif not address:
            messages.error(request, 'Physical Address is required.')
        elif Building.objects.filter(building_id=building_id, deleted=False).exists():
            messages.error(request, 'A building with this Building ID already exists.')
        elif Building.objects.filter(name=name, deleted=False).exists():
            messages.error(request, 'A building with this name already exists.')
        else:
            Building.objects.create(
                building_id=building_id,
                name=name,
                surname=surname,
                division_id=division_id,
                subdivision_id=subdivision_id,
                section_id=section_id,
                address=address,
                latitude=latitude,
                longitude=longitude,
                status=status,
                created_by=request.user,
            )
            messages.success(request, 'Building added successfully.')
            return redirect('backend:building_list')

    context = {'action': 'Add', 'divisions': divisions, 'subdivisions': subdivisions, 'sections': sections}
    return render(request, 'building/add.html', context)

@login_required
def building_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/location/building/'):
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Building, pk=data_id, deleted=False)
    divisions = Division.objects.filter(deleted=False, is_active=True)
    subdivisions = SubDivision.objects.filter(deleted=False, is_active=True)
    sections = Section.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        surname = request.POST.get('surname', '').strip()
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

        if not name:
            messages.error(request, 'Building Name is required.')
        elif not section_id:
            messages.error(request, 'Section is required.')
        elif not address:
            messages.error(request, 'Physical Address is required.')
        elif Building.objects.filter(name=name, deleted=False).exclude(pk=data_id).exists():
            messages.error(request, 'A building with this name already exists.')
        else:
            obj.name = name
            obj.surname = surname
            obj.division_id = division_id
            obj.subdivision_id = subdivision_id
            obj.section_id = section_id
            obj.address = address
            obj.latitude = latitude
            obj.longitude = longitude
            obj.status = status
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Building updated successfully.')
            return redirect('backend:building_list')

    context = {'action': 'Update', 'obj': obj, 'divisions': divisions, 'subdivisions': subdivisions, 'sections': sections}
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

    building = get_object_or_404(Building, pk=data_id, deleted=False)
    
    # Fetch equipment list for this building
    equipment_list = Equipment.objects.filter(building=building, deleted=False).select_related('equipment_type')
    
    # Calculate counts
    total_assets = equipment_list.count()
    
    # Filter lifts, ac units, substations
    lifts_count = equipment_list.filter(equipment_type__name__icontains='lift').count()
    ac_count = equipment_list.filter(Q(equipment_type__name__icontains='ac') | Q(equipment_type__name__icontains='air conditioner')).count()
    substations_count = equipment_list.filter(Q(equipment_type__name__icontains='substation') | Q(equipment_type__name__icontains='sub-station') | Q(equipment_type__name__icontains='sub station') | Q(equipment_type__name__icontains='sub-stations') | Q(equipment_type__name__icontains='sub-stns')).count()
    
    context = {
        'building': building,
        'equipment_list': equipment_list,
        'total_assets': total_assets,
        'lifts_count': lifts_count,
        'ac_count': ac_count,
        'substations_count': substations_count,
    }
    return render(request, 'building/details.html', context)


@login_required
def dash_board(request):
    today = date.today()
    seven_days = today + timedelta(days=7)

    # --- Building stats ---
    buildings_qs = Building.objects.filter(deleted=False)
    total_buildings = buildings_qs.count()
    active_buildings = buildings_qs.filter(status='Active').count()
    inactive_buildings = buildings_qs.filter(status='Inactive').count()
    renovation_buildings = buildings_qs.filter(status='Under Maintenance').count()

    # --- Equipment stats ---
    equipment_qs = Equipment.objects.filter(deleted=False)
    total_equipment = equipment_qs.count()
    eq_active = equipment_qs.filter(status='active').count()
    eq_maintenance = equipment_qs.filter(status='under_maintenance').count()
    eq_repair = equipment_qs.filter(status='in_repair').count()
    eq_outdated = equipment_qs.filter(status='inactive').count()
    eq_decommissioned = equipment_qs.filter(status='decommissioned').count()

    # Percentages for donut chart
    if total_equipment > 0:
        uptime_pct = round(eq_active / total_equipment * 100, 1)
        repair_pct = round(eq_repair / total_equipment * 100, 1)
        outdated_pct = round(eq_outdated / total_equipment * 100, 1)
    else:
        uptime_pct = repair_pct = outdated_pct = 0

    # --- Tickets ---
    tickets_qs = Ticket.objects.filter(deleted=False)
    open_tickets = tickets_qs.filter(status__in=['open', 'assigned', 'in_progress'])
    open_ticket_count = open_tickets.count()
    critical_ticket_count = open_tickets.filter(priority='critical').count()
    assigned_ticket_count = open_tickets.filter(status__in=['assigned', 'in_progress']).count()
    unassigned_ticket_count = open_tickets.filter(status='open').count()
    recent_tickets = (
        open_tickets
        .select_related('equipment', 'building', 'equipment__equipment_type')
        .order_by('-opened_at')[:5]
    )

    # --- Upcoming maintenance (next 7 days) ---
    upcoming_maintenance = (
        ScheduledMaintenance.objects
        .filter(deleted=False, status='pending', scheduled_date__gte=today, scheduled_date__lte=seven_days)
        .select_related('equipment', 'equipment__building')
        .order_by('scheduled_date')
    )
    upcoming_maintenance_count = upcoming_maintenance.count()
    overdue_maintenance = (
        ScheduledMaintenance.objects
        .filter(deleted=False, status='pending', scheduled_date__lt=today)
    )
    overdue_count = overdue_maintenance.count()

    # --- Critical alerts ---
    alerts_qs = CriticalAlert.objects.filter(is_acknowledged=False).order_by('-id')
    new_alert_count = alerts_qs.count()
    critical_alerts = alerts_qs[:5]

    context = {
        # Buildings
        'total_buildings': total_buildings,
        'active_buildings': active_buildings,
        'inactive_buildings': inactive_buildings,
        'renovation_buildings': renovation_buildings,
        # Equipment
        'total_equipment': total_equipment,
        'eq_active': eq_active,
        'eq_maintenance': eq_maintenance,
        'eq_repair': eq_repair,
        'eq_outdated': eq_outdated,
        'eq_decommissioned': eq_decommissioned,
        'uptime_pct': uptime_pct,
        'repair_pct': repair_pct,
        'outdated_pct': outdated_pct,
        # Tickets
        'open_ticket_count': open_ticket_count,
        'critical_ticket_count': critical_ticket_count,
        'assigned_ticket_count': assigned_ticket_count,
        'unassigned_ticket_count': unassigned_ticket_count,
        'recent_tickets': recent_tickets,
        # Maintenance
        'upcoming_maintenance_count': upcoming_maintenance_count,
        'overdue_count': overdue_count,
        # Alerts
        'new_alert_count': new_alert_count,
        'critical_alerts': critical_alerts,
    }
    return render(request, 'dashboard.html', context)




@login_required
def technician_list(request):
    if not checkUserPermission(request, 'can_view', '/backend/technician/'):
        messages.error(request, 'You do not have permission to view this page.')
        return render(request, '403.html', status=403)
    
    technicians = Technician.objects.filter(deleted=False)
    return render(request, 'technician/list.html', {'technicians': technicians}) 


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
            messages.success(request, 'Technician added successfully.')
            return redirect('backend:technician_list')
     
    users = User.objects.filter(is_active=True, technician_profile__isnull=True)
    context = {
        'users': users
    }
    
    return render(request, 'technician/add.html', context)


@login_required
def technician_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/technician/'):
        messages.error(request, 'You do not have permission to edit this technician.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Technician, pk=data_id, deleted=False)
    
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
        messages.success(request, 'Technician updated successfully.')
        return redirect('backend:technician_list')

    users = User.objects.filter(is_active=True, technician_profile__isnull=True)
    context = {
        'obj': obj,
        'users': users
    }
    return render(request, 'technician/update.html', context)


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
    
    equipment_types = EquipmentType.objects.filter(deleted=False)
    return render(request, 'equipment_type/list.html', {'equipment_types': equipment_types})


@login_required
def equipment_type_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/equipment/equipment_type/'):
        messages.error(request, 'You do not have permission to add this equipment type.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        description = request.POST.get('description', '').strip() 
        
        if not name:
            messages.error(request, 'Name is required.')
        else:
            obj = EquipmentType(
                name=name,
                description=description, 
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
    
    equipment_type_datas = EquipmentTypeData.objects.filter(deleted=False)
    return render(request, 'equipment_type_data/list.html', {'equipment_type_datas': equipment_type_datas})


@login_required
def equipment_type_data_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/equipment/equipment_type_data/'):
        messages.error(request, 'You do not have permission to add this equipment type data.')
        return render(request, '403.html', status=403)
    
    if request.method == 'POST':
        equipment_type_id = request.POST.get('equipment_type_id', '').strip()
        field_name = request.POST.get('field_name', '').strip()
        field_type = request.POST.get('field_type', '').strip()
        is_required = request.POST.get('is_required', 'False')
    
        
        if not equipment_type_id or not field_name or not field_type:
            messages.error(request, 'Equipment Type, Field Name, and Field Type are required.')
        else:
            obj = EquipmentTypeData(
                equipment_type_id=equipment_type_id,
                field_name=field_name,
                field_type=field_type,
                is_required=is_required,
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
        is_required = request.POST.get('is_required', 'False')
      
        if not equipment_type_id or not field_name or not field_type:
            messages.error(request, 'Equipment Type, Field Name, and Field Type are required.')
        else:
            obj.equipment_type_id = equipment_type_id
            obj.field_name = field_name
            obj.field_type = field_type
            obj.is_required = is_required
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
    
    qs = EquipmentComponentPreset.objects.filter(deleted=False).select_related('equipment_type')
    if search:
        qs = qs.filter(name__icontains=search)
    if equipment_type_id:
        qs = qs.filter(equipment_type_id=equipment_type_id)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    
    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'search': search,
        'equipment_type_id': equipment_type_id,
        'equipment_types': equipment_types,
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
        
        if not name or not equipment_type_id:
            messages.error(request, 'Name and Equipment Type are required.')
        else:
            EquipmentComponentPreset.objects.create(
                name=name,
                equipment_type_id=equipment_type_id,
                description=description,
                created_by=request.user,
                is_active=True
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
    
    qs = Equipment.objects.filter(deleted=False).select_related('equipment_type', 'building').annotate(components_count=Count('components'))
    
    if search:
        qs = qs.filter(Q(equipment_id__icontains=search) | Q(brand__icontains=search) | Q(building__name__icontains=search))
    if type_id:
        qs = qs.filter(equipment_type_id=type_id)
    if status:
        qs = qs.filter(status=status)
        
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)
    
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    
    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'search': search,
        'type_id': type_id,
        'status': status,
        'equipment_types': equipment_types,
        'status_choices': Equipment.STATUS_CHOICES,
    }   
    return render(request, 'equipment/list.html', context)

from backend.models import EquipmentComponents 

@login_required
def equipment_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to add equipment.')
        return render(request, '403.html', status=403)
        
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    buildings = Building.objects.filter(deleted=False, is_active=True)
    presets = EquipmentComponentPreset.objects.filter(deleted=False, is_active=True)
    
    if request.method == 'POST':
        equipment_id = request.POST.get('equipment_id', '').strip()
        equipment_type_id = request.POST.get('equipment_type_id', '')
        building_id = request.POST.get('building_id', '')
        brand = request.POST.get('brand', '').strip()
        floor_location = request.POST.get('floor_location', '').strip()

        # presets
        component_preset_ids = request.POST.getlist('component_preset')

        # components 
        component_names = request.POST.getlist('component_name[]')
        component_descriptions = request.POST.getlist('component_description[]')
        
        maintenance_period_days = request.POST.get('maintenance_period_days', '90')
        next_service_due = request.POST.get('next_service_due', '')
        
        if not equipment_id or not equipment_type_id or not building_id:
            messages.error(request, 'Equipment ID, Type, and Building are required.')
        else:
            obj = Equipment(
                equipment_id=equipment_id,
                equipment_type_id=equipment_type_id,
                building_id=building_id,
                brand=brand,
                floor_location=floor_location,
                maintenance_period_days=maintenance_period_days if maintenance_period_days else 90,
                next_service_due=next_service_due if next_service_due else None,
                created_by=request.user
            )
            obj.save()
            
            # Save presets relation
            if component_preset_ids:
                obj.component_preset.set(component_preset_ids)
            
            # Save components
            for i in range(len(component_names)):
                c_name = component_names[i].strip()
                c_desc = component_descriptions[i].strip() if i < len(component_descriptions) else ''
                if c_name:
                    EquipmentComponents.objects.create(
                        equipment=obj,
                        name=c_name,
                        description=c_desc,
                        created_by=request.user
                    )
            
            messages.success(request, 'Equipment added successfully.')
            return redirect('backend:equipment_list')
    
    context = {
        'action': 'Add',
        'equipment_types': equipment_types,
        'buildings': buildings,
        'presets': presets,
    }
    return render(request, 'equipment/add.html', context)



@login_required
def equipment_update(request, data_id):
    if not checkUserPermission(request, 'can_edit', '/backend/equipment/equipment/'):
        messages.error(request, 'You do not have permission to edit this equipment.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Equipment, pk=data_id, deleted=False)
    equipment_types = EquipmentType.objects.filter(deleted=False, is_active=True)
    buildings = Building.objects.filter(deleted=False, is_active=True)
    presets = EquipmentComponentPreset.objects.filter(deleted=False, is_active=True)
    
    if request.method == 'POST':
        equipment_id = request.POST.get('equipment_id', '').strip()
        equipment_type_id = request.POST.get('equipment_type_id', '')
        building_id = request.POST.get('building_id', '')
        brand = request.POST.get('brand', '').strip()
        floor_location = request.POST.get('floor_location', '').strip()
        
        # presets
        component_preset_ids = request.POST.getlist('component_preset')

        # components
        component_names = request.POST.getlist('component_name[]')
        component_descriptions = request.POST.getlist('component_description[]')
        
        maintenance_period_days = request.POST.get('maintenance_period_days', '90')
        next_service_due = request.POST.get('next_service_due', '')
        
        status = request.POST.get('status', obj.status)
        
        if not equipment_id or not equipment_type_id or not building_id:
            messages.error(request, 'Equipment ID, Type, and Building are required.')
        else:
            obj.equipment_id = equipment_id
            obj.equipment_type_id = equipment_type_id
            obj.building_id = building_id
            obj.brand = brand
            obj.floor_location = floor_location
            obj.maintenance_period_days = maintenance_period_days if maintenance_period_days else 90
            obj.next_service_due = next_service_due if next_service_due else None
            obj.status = status
            
            obj.updated_by = request.user
            obj.save()
            
            # Save presets relation
            obj.component_preset.set(component_preset_ids)

            # Recreate components
            obj.components.all().delete()
            for i in range(len(component_names)):
                c_name = component_names[i].strip()
                c_desc = component_descriptions[i].strip() if i < len(component_descriptions) else ''
                if c_name:
                    EquipmentComponents.objects.create(
                        equipment=obj,
                        name=c_name,
                        description=c_desc,
                        created_by=request.user
                    )
                    
            messages.success(request, 'Equipment updated successfully.')
            return redirect('backend:equipment_list')
    
    context = {
        'action': 'Update',
        'obj': obj,
        'equipment_types': equipment_types,
        'buildings': buildings,
        'presets': presets,
    }
    return render(request, 'equipment/update.html', context)


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

@login_required
def qr_equipement_views(request, data_id):
    obj = get_object_or_404(Equipment, pk=data_id, deleted=False)
    
    if request.GET.get('view') == '1':
        last_maintenance = obj.maintenance_records.filter(deleted=False).order_by('-maintenance_date').first()
        
        add_url = request.build_absolute_uri(reverse('backend:static_otp_verification'))
        qr_data = f"{add_url}?building={obj.building.id}&equipment={obj.id}"
        img = generate_qr_code(qr_data)
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        qr_base64 = base64.b64encode(buffer.getvalue()).decode()

        context = {
            'obj': obj,
            'last_maintenance': last_maintenance,
            'qr_base64': qr_base64,
        }
        return render(request, "equipment/qr.html", context)
    
    # Redirect to static_otp_verification with auto-fill parameters
    redirect_url = reverse('backend:static_otp_verification')
    return redirect(f"{redirect_url}?building={obj.building.id}&equipment={obj.id}")

# ======================================== Maintenance Record Review ========================================

@login_required
def maintenance_record_review(request, record_id):
    """Approve or Reject a MaintenanceRecord from the equipment detail page."""
    from django.utils import timezone as tz

    record = get_object_or_404(MaintenanceRecord, pk=record_id, deleted=False)
    referer = request.META.get('HTTP_REFERER')
    fallback_url = reverse('backend:equipment_detail', kwargs={'data_id': record.equipment_id})
    redirect_url = referer if referer else fallback_url

    if request.method != 'POST':
        messages.error(request, 'Invalid request method.')
        return redirect(redirect_url)

    action = request.POST.get('action', '').strip()   # 'approve' or 'reject'
    review_notes = request.POST.get('review_notes', '').strip()

    if action == 'approve':
        record.review_status = 'approved'
        record.review_notes = review_notes
        record.reviewed_by = request.user
        record.reviewed_at = tz.now()
        record.updated_by = request.user
        record.save()
        messages.success(request, 'Maintenance record approved successfully.')

    elif action == 'reject':
        if not review_notes:
            messages.error(request, 'Please provide a reason for rejection.')
            return redirect(redirect_url)
        record.review_status = 'rejected'
        record.review_notes = review_notes
        record.reviewed_by = request.user
        record.reviewed_at = tz.now()
        record.updated_by = request.user
        record.save()
        messages.success(request, 'Maintenance record rejected.')

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
    equipment_id  = request.GET.get('equipment_id', '').strip()
    record_type   = request.GET.get('record_type', '').strip()
    review_status = request.GET.get('review_status', '').strip()
    date_from     = request.GET.get('date_from', '').strip()
    date_to       = request.GET.get('date_to', '').strip()

    qs = MaintenanceRecord.objects.filter(deleted=False).select_related(
        'equipment', 'equipment__building', 'equipment__equipment_type',
        'technician', 'reviewed_by'
    ).order_by('-maintenance_date', '-created_at')

    if search:
        qs = qs.filter(
            Q(equipment__equipment_id__icontains=search)
            | Q(equipment__building__name__icontains=search)
            | Q(work_description__icontains=search)
            | Q(technician__first_name__icontains=search)
            | Q(technician__last_name__icontains=search)
            | Q(technician__username__icontains=search)
        )
    if equipment_id:
        qs = qs.filter(equipment_id=equipment_id)
    if record_type:
        qs = qs.filter(record_type=record_type)
    if review_status:
        qs = qs.filter(review_status=review_status)
    if date_from:
        try:
            from datetime import datetime
            qs = qs.filter(maintenance_date__gte=datetime.strptime(date_from, '%Y-%m-%d').date())
        except ValueError:
            pass
    if date_to:
        try:
            from datetime import datetime
            qs = qs.filter(maintenance_date__lte=datetime.strptime(date_to, '%Y-%m-%d').date())
        except ValueError:
            pass

    # ── Summary stats (from the FULL unfiltered set) ──────────────────────
    all_records    = MaintenanceRecord.objects.filter(deleted=False)
    total_count    = all_records.count()
    pending_count  = all_records.filter(review_status='pending').count()
    approved_count = all_records.filter(review_status='approved').count()
    rejected_count = all_records.filter(review_status='rejected').count()

    # ── Pagination ────────────────────────────────────────────────────────
    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)

    # ── Filter options ────────────────────────────────────────────────────
    equipment_list = Equipment.objects.filter(deleted=False, is_active=True).order_by('equipment_id')

    context = {
        'data_list'     : data_list,
        'paginator_list': paginator_list,
        'last_page'     : last_page,
        # active filters
        'search'        : search,
        'f_equipment_id': equipment_id,
        'f_record_type' : record_type,
        'f_review_status': review_status,
        'f_date_from'   : date_from,
        'f_date_to'     : date_to,
        # filter options
        'equipment_list': equipment_list,
        'record_type_choices': MaintenanceRecord.RECORD_TYPE_CHOICES,
        'review_status_choices': MaintenanceRecord.REVIEW_STATUS_CHOICES,
        # stats
        'total_count'   : total_count,
        'pending_count' : pending_count,
        'approved_count': approved_count,
        'rejected_count': rejected_count,
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
    
    # Optional filtering
    status_filter = request.GET.get('status', '')
    priority_filter = request.GET.get('priority', '')
    
    tickets = Ticket.objects.filter(deleted=False).order_by('-created_at')
    
    if status_filter:
        tickets = tickets.filter(status=status_filter)
    if priority_filter:
        tickets = tickets.filter(priority=priority_filter)
        
    context = {
        'tickets': tickets,
        'status_filter': status_filter,
        'priority_filter': priority_filter,
    }
    return render(request, 'ticket/list.html', context)

@login_required
def ticket_detail(request, data_id):
    if not checkUserPermission(request, 'can_view', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to view this ticket.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Ticket, pk=data_id, deleted=False)

    if request.method == 'POST':
        if not checkUserPermission(request, 'can_update', '/backend/ticket/ticket/'):
            messages.error(request, 'You do not have permission to update this ticket.')
            return redirect('backend:ticket_detail', data_id=obj.id)

        # --- Handle comment submission ---
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
            return redirect('backend:ticket_detail', data_id=obj.id)

        # --- Handle technician assignment ---
        technician_id = request.POST.get('technician_id')
        if technician_id:
            try:
                technician_obj = Technician.objects.get(id=technician_id)
                
                # Check if already assigned
                if not AssignTechnician.objects.filter(ticket=obj, technician=technician_obj, deleted=False).exists():
                    AssignTechnician.objects.create(
                        ticket=obj,
                        technician=technician_obj,
                        assign_by=request.user
                    )

                    # Also create an AssignActivities record so it appears on the technician's task list
                    if not AssignActivities.objects.filter(ticket=obj, technician=technician_obj, is_active=True).exists():
                        AssignActivities.objects.create(
                            ticket=obj,
                            technician=technician_obj,
                            status='pending',
                            completion_status='pending',
                            created_by=request.user,
                        )
                    
                    if obj.status == 'open':
                        obj.status = 'assigned'
                        obj.assigned_at = timezone.now()
                        obj.save()
                        
                    TicketActivityLog.objects.create(
                        ticket=obj,
                        actor=request.user,
                        action_type='assigned',
                        description=f"Assigned technician: {technician_obj.display_name}"
                    )
                    
                    messages.success(request, 'Technician assigned successfully.')
                else:
                    messages.warning(request, 'Technician is already assigned to this ticket.')
            except Technician.DoesNotExist:
                messages.error(request, 'Selected technician does not exist.')
                
        return redirect('backend:ticket_detail', data_id=obj.id)
        
    technicians = Technician.objects.filter(deleted=False, is_active=True)
    
    context = {
        'obj': obj,
        'technicians': technicians,
    }
    return render(request, 'ticket/detail.html', context) 

@login_required
def ticket_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to add this ticket.')
        return render(request, '403.html', status=403)

    equipments = Equipment.objects.filter(deleted=False, is_active=True)
    buildings = Building.objects.filter(deleted=False, is_active=True)

    if request.method == 'POST':
        title = request.POST.get('title', '').strip()
        equipment_id = request.POST.get('equipment_id', '')
        building_id = request.POST.get('building_id', '')
        issue_type = request.POST.get('issue_type', 'breakdown')
        priority = request.POST.get('priority', 'medium')
        status = request.POST.get('status', 'open')
        issue_description = request.POST.get('issue_description', '').strip()
        
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
                user=request.user,
                created_by=request.user
            )
            obj.save()
            messages.success(request, 'Ticket created successfully.')
            return redirect('backend:ticket_list')
    
    # Pre-fill equipment and building from GET parameters if provided
    initial_eq = request.GET.get('equipment_id') or request.GET.get('equipment')
    initial_b = request.GET.get('building_id') or request.GET.get('building')
    
    obj = None
    if initial_eq or initial_b:
        class DummyObj:
            def __init__(self, eq_id, b_id):
                try:
                    self.equipment_id = int(eq_id) if eq_id else None
                except ValueError:
                    self.equipment_id = None
                try:
                    self.building_id = int(b_id) if b_id else None
                except ValueError:
                    self.building_id = None
                self.status = 'open'
                self.issue_type = 'breakdown'
        obj = DummyObj(initial_eq, initial_b)

    context = {
        'action': 'Add',
        'equipments': equipments,
        'buildings': buildings,
        'obj': obj,
    }
    return render(request, 'ticket/add.html', context)

@login_required
def ticket_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/ticket/ticket/'):
        messages.error(request, 'You do not have permission to update this ticket.')
        return render(request, '403.html', status=403)
    
    obj = get_object_or_404(Ticket, pk=data_id, deleted=False)
    equipments = Equipment.objects.filter(deleted=False, is_active=True)
    buildings = Building.objects.filter(deleted=False, is_active=True)
    
    if request.method == 'POST':
        title = request.POST.get('title', '').strip()
        equipment_id = request.POST.get('equipment_id', '')
        building_id = request.POST.get('building_id', '')
        issue_type = request.POST.get('issue_type', 'breakdown')
        priority = request.POST.get('priority', 'medium')
        status = request.POST.get('status', 'open')
        issue_description = request.POST.get('issue_description', '').strip()
        
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
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Ticket updated successfully.')
            return redirect('backend:ticket_list')
    
    context = {
        'action': 'Update',
        'obj': obj,
        'equipments': equipments,
        'buildings': buildings,
    }
    return render(request, 'ticket/update.html', context)


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
    qs = Maintenance.objects.filter(deleted=False).prefetch_related(
        'equipment',
        'equipment__equipment_type'
    )

    if search:
        qs = qs.filter(
            Q(maintenance_serial__icontains=search) |
            Q(qr_code__icontains=search)
        )

    qs = qs.annotate(
        num_records=Count(
            'equipment__maintenance_records',
            filter=Q(equipment__maintenance_records__deleted=False),
            distinct=True
        )
    ).order_by('-created_at')

    page_num = request.GET.get('page', 1)
    data_list, paginator_list, last_page = paginate_data(request, page_num, qs)

    context = {
        'data_list': data_list,
        'paginator_list': paginator_list,
        'last_page': last_page,
        'search': search,
    }
    return render(request, 'maintenance/list.html', context)


@login_required
def maintenance_add(request):
    if not checkUserPermission(request, 'can_add', '/backend/maintenance/'):
        messages.error(request, 'You do not have permission to add a maintenance record.')
        return render(request, '403.html', status=403)

    buildings = Building.objects.filter(deleted=False, is_active=True)
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
                components_data.append({
                    'name': comp_name,
                    'is_checked': is_checked,
                    'remark': remark,
                    'suggestion': suggestion
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
def maintenance_update(request, data_id):
    if not checkUserPermission(request, 'can_update', '/backend/maintenance/'):
        messages.error(request, 'You do not have permission to edit this record.')
        return render(request, '403.html', status=403)

    obj = get_object_or_404(Maintenance, pk=data_id, deleted=False)
    buildings = Building.objects.filter(deleted=False, is_active=True)
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
                components_data.append({
                    'name': comp_name,
                    'is_checked': is_checked,
                    'remark': remark,
                    'suggestion': suggestion
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
        ).values('id', 'equipment_id').order_by('equipment_id')
        
        equipment = [
            {
                'id': item['id'],
                'equipment_id': item['equipment_id'],
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


@login_required
def static_otp_verification(request):
    building_id = request.GET.get('building')
    equipment_id = request.GET.get('equipment')
    context = {
        'building_id': building_id,
        'equipment_id': equipment_id,
    }
    return render(request, 'maintenance/static_otp.html', context) 


