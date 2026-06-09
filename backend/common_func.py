from threading import Thread
from django.conf import settings
from django.template.loader import render_to_string
from django.core.mail import  get_connection

from backend.models import UserMenuPermission


def checkUserPermission(request, access_type, menu_url):
    try:
        user_permissions = {
            "can_view": "can_view",
            "can_add": "can_add",
            "can_update": "can_update",
            "can_delete": "can_delete",
        }

        if request.user.is_superuser:
            return True

        check_user_permission = UserMenuPermission.objects.filter(
            user_id=request.user.id, is_active=True, **{user_permissions[access_type]: True}, menu__menu_url=menu_url,
        )

        if check_user_permission:
            return True
        else:
            return False
    except Exception:
        return False
