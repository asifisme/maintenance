from django.contrib import admin
from .models import (
    WebImages, PasswordResetCode, LoginLog, BackendMenu, UserMenuPermission, SiteSettings,
    Division, SubDivision, Section, Building, Technician, EquipmentType, EquipmentTypeData,
    Equipment, MaintenanceRecord, MaintenanceAttachment, Ticket, TicketComment,
    TicketActivityLog, ScheduledMaintenance, CriticalAlert
)

@admin.register(WebImages)
class WebImagesAdmin(admin.ModelAdmin):
    list_display = ("unique_key", "created_at", "created_by")

@admin.register(PasswordResetCode)
class PasswordResetCodeAdmin(admin.ModelAdmin):
    list_display = ("user", "code", "is_used", "created_at", "expires_at")
    list_filter = ("is_used",)

@admin.register(LoginLog)
class LoginLogAdmin(admin.ModelAdmin):
    list_display = ("username", "login_status", "login_ip", "created_at")
    list_filter = ("login_status",)

@admin.register(BackendMenu)
class BackendMenuAdmin(admin.ModelAdmin):
    list_display = ("menu_name", "module_name", "is_main_menu", "is_active")
    list_filter = ("module_name", "is_active", "is_main_menu")

@admin.register(UserMenuPermission)
class UserMenuPermissionAdmin(admin.ModelAdmin):
    list_display = ("user", "menu", "can_view", "can_add", "can_update", "can_delete", "is_active")
    list_filter = ("is_active", "can_view")

@admin.register(SiteSettings)
class SiteSettingsAdmin(admin.ModelAdmin):
    list_display = ("site_title", "contact_email", "contact_phone", "is_active")

@admin.register(Division)
class DivisionAdmin(admin.ModelAdmin):
    list_display = ("name", "code", "is_active", "created_at")

@admin.register(SubDivision)
class SubDivisionAdmin(admin.ModelAdmin):
    list_display = ("name", "code", "division", "is_active")
    list_filter = ("division", "is_active")

@admin.register(Section)
class SectionAdmin(admin.ModelAdmin):
    list_display = ("name", "code", "subdivision", "is_active")
    list_filter = ("subdivision", "is_active")

@admin.register(Building)
class BuildingAdmin(admin.ModelAdmin):
    list_display = ("building_id", "name", "surname", "division", "status", "is_active")
    list_filter = ("status", "is_active", "division")

@admin.register(Technician)
class TechnicianAdmin(admin.ModelAdmin):
    list_display = ("user", "phone_number", "is_active")

@admin.register(EquipmentType)
class EquipmentTypeAdmin(admin.ModelAdmin):
    list_display = ("name", "is_active", "created_at")

@admin.register(EquipmentTypeData)
class EquipmentTypeDataAdmin(admin.ModelAdmin):
    list_display = ("equipment_type", "field_name", "field_type", "is_required")
    list_filter = ("equipment_type", "field_type")

@admin.register(Equipment)
class EquipmentAdmin(admin.ModelAdmin):
    list_display = ("equipment_id", "equipment_type", "building", "brand", "status")
    list_filter = ("status", "equipment_type", "building")
    search_fields = ("equipment_id", "brand", "model_number", "serial_number")

@admin.register(MaintenanceRecord)
class MaintenanceRecordAdmin(admin.ModelAdmin):
    list_display = ("equipment", "record_type", "maintenance_date", "technician", "review_status")
    list_filter = ("record_type", "review_status", "maintenance_date")

@admin.register(MaintenanceAttachment)
class MaintenanceAttachmentAdmin(admin.ModelAdmin):
    list_display = ("record", "uploaded_by", "caption")

@admin.register(Ticket)
class TicketAdmin(admin.ModelAdmin):
    list_display = ("title", "status", "priority", "issue_type", "equipment", "opened_at")
    list_filter = ("status", "priority", "issue_type", "opened_at")
    search_fields = ("title",)

@admin.register(TicketComment)
class TicketCommentAdmin(admin.ModelAdmin):
    list_display = ("ticket", "author", "is_internal", "created_at")
    list_filter = ("is_internal",)

@admin.register(TicketActivityLog)
class TicketActivityLogAdmin(admin.ModelAdmin):
    list_display = ("ticket", "actor", "action_type")
    list_filter = ("action_type",)

@admin.register(ScheduledMaintenance)
class ScheduledMaintenanceAdmin(admin.ModelAdmin):
    list_display = ("title", "equipment", "scheduled_date", "recurrence", "status")
    list_filter = ("status", "recurrence", "scheduled_date")

@admin.register(CriticalAlert)
class CriticalAlertAdmin(admin.ModelAdmin):
    list_display = ("title", "alert_type", "severity", "equipment", "is_acknowledged")
    list_filter = ("severity", "alert_type", "is_acknowledged")
