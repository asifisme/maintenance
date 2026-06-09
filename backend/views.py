import code
import os
import base64
import logging
import json
import calendar
from datetime import datetime, date 
from urllib import request
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
    Division, SubDivision
)

from backend.forms import (
    CustomUserLoginForm,
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
        is_active = request.POST.get('is_active') == '1'

        if not name or not code:
            messages.error(request, 'Name and Code are required.')
        elif Division.objects.filter(code=code, deleted=False).exists():
            messages.error(request, 'A division with this code already exists.')
        else:
            Division.objects.create(
                name=name, code=code, location=location or None,
                is_active=is_active, created_by=request.user,
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
        is_active = request.POST.get('is_active') == '1'

        if not name or not code:
            messages.error(request, 'Name and Code are required.')
        elif Division.objects.filter(code=code, deleted=False).exclude(pk=data_id).exists():
            messages.error(request, 'A division with this code already exists.')
        else:
            obj.name = name
            obj.code = code
            obj.location = location or None
            obj.is_active = is_active
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Division updated successfully.')
            return redirect('backend:division_list')

    context = {'action': 'Update', 'obj': obj}
    return render(request, 'division/update.html', context)


@login_required
def division_delete(request, data_id):
    if not checkUserPermission(request, "can_delete", "/backend/location/division/"):
        messages.error(request, 'You do not have permission to delete this division.')
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        Division.objects.filter(pk=data_id, deleted=False).update(
            deleted=True,
            is_active=False,
            updated_by=request.user
        )
        messages.success(request, 'Division deleted successfully.')
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
        location = request.POST.get('location', '').strip()
        is_active = request.POST.get('is_active') == '1'

        if not division_id or not name:
            messages.error(request, 'Division and Name are required.')
        else:
            SubDivision.objects.create(
                division_id=division_id, name=name, location=location or None,
                is_active=is_active, created_by=request.user,
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
        location = request.POST.get('location', '').strip()
        is_active = request.POST.get('is_active') == '1'

        if not division_id or not name:
            messages.error(request, 'Division and Name are required.')
        else:
            obj.division_id = division_id
            obj.name = name
            obj.location = location or None
            obj.is_active = is_active
            obj.updated_by = request.user
            obj.save()
            messages.success(request, 'Sub-Division updated successfully.')
            return redirect('backend:subdivision_list')

    context = {'action': 'Update', 'obj': obj, 'divisions': divisions}
    return render(request, 'subdivision/update.html', context)


@login_required
def subdivision_delete(request, data_id):
    if not checkUserPermission(request, "can_delete", "/backend/location/subdivision/"):
        messages.error(request, 'You do not have permission to delete this sub-division.')
        return render(request, "403.html", status=403)

    if request.method == 'POST':
        SubDivision.objects.filter(pk=data_id, deleted=False).update(
            deleted=True,
            is_active=False,
            updated_by=request.user
        )
        messages.success(request, 'Sub-Division deleted successfully.')
        return redirect('backend:subdivision_list')
    
    messages.error(request, 'Invalid request method.')
    return redirect('backend:subdivision_list')
