from django.db.models import Q
from backend.models import (
    BackendMenu, UserMenuPermission, SiteSettings, Notification
)


def menu_items(request):
    current_path = request.get_full_path()

    breadcrumb_menu = []
    breadcrumb = None
    best_match_length = 0

    dashboard_active = True

    menu_list = list(
        UserMenuPermission.objects.filter(
            user_id=request.user.id, menu__is_main_menu=True, can_view=True, menu__parent__isnull=True,
            menu__is_active=True, is_active=True,
        )
        .select_related('menu', 'user')
        .order_by('menu__id')
    )

    search_menu_list = (
        UserMenuPermission.objects.filter(
            user_id=request.user.id, can_view=True, menu__is_active=True, is_active=True,
        )
        .select_related('menu', 'user')
        .order_by('menu__id')
    )

    # Find best match and track hierarchy
    best_item = None
    best_match_length = 0

    for menu in menu_list:
        menu.is_current = False
        m_url = (menu.menu.menu_url or "").strip()
        if m_url and m_url in current_path:
            m_len = len(m_url)
            if m_len > best_match_length:
                best_match_length = m_len
                best_item = (menu, None, None)

        sub_menu = list(
            UserMenuPermission.objects.filter(
                user_id=request.user.id, menu__is_sub_menu=True, can_view=True, menu__parent_id=menu.menu.id,
                menu__is_active=True, is_active=True,
            )
            .select_related('menu', 'user')
            .order_by('menu__id')
        )

        for sub in sub_menu:
            sub.is_current = False
            s_url = (sub.menu.menu_url or "").strip()
            if s_url and s_url in current_path:
                s_len = len(s_url)
                if s_len > best_match_length:
                    best_match_length = s_len
                    best_item = (menu, sub, None)

            child_menu = list(
                UserMenuPermission.objects.filter(
                    user_id=request.user.id, menu__is_sub_child_menu=True, can_view=True, menu__parent_id=sub.menu.id,
                    menu__is_active=True, is_active=True,
                )
                .select_related('menu', 'user')
                .order_by('menu__id')
            )

            for child in child_menu:
                child.is_current = False
                c_url = (child.menu.menu_url or "").strip()
                if c_url and c_url in current_path:
                    c_len = len(c_url)
                    if c_len > best_match_length:
                        best_match_length = c_len
                        best_item = (menu, sub, child)

            sub.child_menu = child_menu

        menu.sub_menu = sub_menu

    if best_item:
        top_m, sub_m, ch_m = best_item
        dashboard_active = False
        if ch_m:
            ch_m.is_current = True
            breadcrumb = ch_m.menu
        if sub_m:
            sub_m.is_current = True
            if not breadcrumb:
                breadcrumb = sub_m.menu
        if top_m:
            top_m.is_current = True
            if not breadcrumb:
                breadcrumb = top_m.menu

    # Separate User Management so sidebar can render it below Help & Support
    sidebar_main_menus = []
    user_management_menu = None
    for menu in menu_list:
        menu_name_lower = (menu.menu.menu_name or "").strip().lower()
        menu_url_str = (menu.menu.menu_url or "").strip().lower()
        if menu_url_str == "/user/" or menu_name_lower == "user management":
            user_management_menu = menu
        else:
            sidebar_main_menus.append(menu)

    # Build breadcrumb trail
    breadcrumb_menu = []
    if breadcrumb:
        breadcrumb_menu.append(breadcrumb)
        while breadcrumb and breadcrumb.parent_id:
            breadcrumb = BackendMenu.objects.filter(id=breadcrumb.parent_id, is_active=True).first()
            if breadcrumb:
                breadcrumb_menu.append(breadcrumb)
        breadcrumb_menu.reverse()

    header_notifications = []
    unread_notifications_count = 0
    if request.user.is_authenticated:
        notif_qs = Notification.objects.filter(
            deleted=False
        ).filter(
            Q(recipient__isnull=True) | Q(recipient=request.user)
        )
        unread_notifications_count = notif_qs.filter(is_read=False).count()
        header_notifications = list(
            notif_qs.select_related('actor', 'ticket', 'equipment', 'technician')[:12]
        )

    return {
        'dashboard_active': dashboard_active,
        "breadcrumb_menu": breadcrumb_menu,
        'main_menu_list': menu_list,
        'sidebar_main_menus': sidebar_main_menus,
        'user_management_menu': user_management_menu,
        'search_menu_list': search_menu_list,
        'header_notifications': header_notifications,
        'unread_notifications_count': unread_notifications_count,
    }


def site_design_settings(request):
    current_path = request.get_full_path()

    if "/admin/" not in current_path:
        site_setting, _ = SiteSettings.objects.get_or_create()

 
        return {
            "site_setting": site_setting,
            "request_path": request.path,
        }
    return {}
