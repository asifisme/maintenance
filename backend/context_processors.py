from backend.models import (
    BackendMenu, UserMenuPermission, SiteSettings
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

    # Find best breadcrumb match
    for menu in menu_list:
        menu.is_current = False
        if menu.menu.menu_url and menu.menu.menu_url in current_path:
            match_length = len(menu.menu.menu_url)
            if match_length > best_match_length:
                best_match_length = match_length
                breadcrumb = menu.menu
            menu.is_current = True
            dashboard_active = False

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
            if sub.menu.menu_url and sub.menu.menu_url in current_path:
                match_length = len(sub.menu.menu_url)
                if match_length > best_match_length:
                    best_match_length = match_length
                    breadcrumb = sub.menu
                sub.is_current = True
                menu.is_current = True
                dashboard_active = False

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
                if child.menu.menu_url and child.menu.menu_url in current_path:
                    match_length = len(child.menu.menu_url)
                    if match_length > best_match_length:
                        best_match_length = match_length
                        breadcrumb = child.menu
                    child.is_current = True
                    sub.is_current = True
                    menu.is_current = True
                    dashboard_active = False

            sub.child_menu = child_menu

        menu.sub_menu = sub_menu

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

    return {
        'dashboard_active': dashboard_active,
        "breadcrumb_menu": breadcrumb_menu,
        'main_menu_list': menu_list,
        'sidebar_main_menus': sidebar_main_menus,
        'user_management_menu': user_management_menu,
        'search_menu_list': search_menu_list,
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
