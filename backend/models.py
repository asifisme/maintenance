import os
import uuid
import string
import random
from datetime import datetime, timedelta
from django.db import models
from django.db.models import JSONField  
from django.contrib.auth.models import User
from django.utils.text import slugify
from django.urls import reverse
from django.db import models
from django.contrib.auth.models import User 
from django.utils import timezone
from io import BytesIO
from django.core.files.base import ContentFile
from django.utils.text import slugify
from django.db import transaction
from django.utils import timezone
from PIL import Image 



def generate_unique_key():
    letters_and_digits = string.ascii_lowercase + string.digits
    while True:
        key = ''.join(random.choices(letters_and_digits, k=32))
        if not WebImages.objects.filter(unique_key=key).exists():
            return key


class WebImages(models.Model):
    unique_key = models.CharField(max_length=100, unique=True, default=generate_unique_key, editable=False)
    image = models.ImageField(upload_to="images/")
    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='web_images')

    @property
    def get_image_url(self):
        if not self.unique_key:
            return None
        path = reverse('site:serve_optimized_image', kwargs={'unique_key': self.unique_key})
        return path

    class Meta:
        db_table = 'web_images'

    def save(self, *args, **kwargs):
        if self.image and not self.image.name.lower().endswith(".webp"):
            img = Image.open(self.image)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")

            base_name = slugify(os.path.splitext(os.path.basename(self.image.name))[0])
            unique_suffix = self.unique_key[:8]

            max_base_length = 90 - len(unique_suffix) - len(".webp") - 1
            if len(base_name) > max_base_length:
                base_name = base_name[:max_base_length]

            safe_filename = f"{base_name}_{unique_suffix}.webp"

            buffer = BytesIO()
            img.save(buffer, format="WEBP", quality=100)
            buffer.seek(0)

            self.image.save(safe_filename, ContentFile(buffer.read()), save=False)

        super().save(*args, **kwargs)

    def __str__(self):
        return self.unique_key


def default_expiry():
    return datetime.now() + timedelta(minutes=5)


class PasswordResetCode(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='password_reset_codes')
    code = models.CharField(max_length=6)
    is_used = models.BooleanField(default=False)
    token = models.UUIDField(default=uuid.uuid4, editable=False, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)
    expires_at = models.DateTimeField(default=default_expiry)

    def __str__(self):
        return f"{self.user.email} - {self.code} - {self.created_at}"

    class Meta:
        db_table = 'password_reset_code'
        verbose_name_plural = 'Password Reset Codes'
        ordering = ['-created_at']


class LoginLog(models.Model):
    user = models.ForeignKey(User, on_delete=models.DO_NOTHING, blank=True, null=True)
    username = models.CharField(max_length=100, blank=True, null=True)
    wrong_password = models.CharField(max_length=100, blank=True, null=True)
    login_ip = models.CharField(max_length=100, blank=True, null=True)
    login_status = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = "login_logs"

    def __str__(self) -> str:
        return self.username or ""


class BackendMenu(models.Model):
    module_name = models.CharField(max_length=100, db_index=True)
    menu_name = models.CharField(max_length=100, db_index=True)
    menu_url = models.CharField(max_length=250, blank=True, null=True)
    menu_icon = models.CharField(max_length=250, blank=True, null=True)
    menu_description = models.TextField(blank=True, null=True)
    parent = models.ForeignKey('self', on_delete=models.CASCADE, related_name='children', blank=True, null=True)
    is_main_menu = models.BooleanField(default=False)
    is_sub_menu = models.BooleanField(default=False)
    is_sub_child_menu = models.BooleanField(default=False)
    is_active = models.BooleanField(default=True)

    class Meta:
        db_table = "backend_menu"
        indexes = [
            models.Index(fields=['module_name', 'is_active']),
        ]

    def __str__(self) -> str:
        return self.menu_name

class UserMenuPermission(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name="user_permission")
    menu = models.ForeignKey(BackendMenu, on_delete=models.CASCADE, related_name="user_permission")
    can_view = models.BooleanField(default=False)
    can_add = models.BooleanField(default=False)
    can_update = models.BooleanField(default=False)
    can_delete = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name="created_by_user_permission")
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name="updated_by_user_permission", blank=True, null=True)
    deleted_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name="deleted_by_user_permission", blank=True, null=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False)

    class Meta:
        db_table = "user_permission"
        unique_together = (('user', 'menu'),)

    def __str__(self):
        return f"{self.user} -> {self.menu}"


class SiteSettings(models.Model):
    site_title = models.CharField(max_length=255, default="Demo HRMS")
    logo = models.ImageField(upload_to='settings/logo/', blank=True, null=True)
    favicon = models.ImageField(upload_to='settings/favicon/', blank=True, null=True)

    contact_email = models.EmailField(blank=True, null=True)
    contact_phone = models.CharField(max_length=20, blank=True, null=True)
    address = models.TextField(blank=True, null=True)

    facebook_url = models.URLField(blank=True, null=True)
    instagram_url = models.URLField(blank=True, null=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='site_settings_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='site_settings_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)


    def __str__(self):
        return self.site_title if self.site_title else "Site Settings"
    

class Division(models.Model):
    name = models.CharField(max_length=120, unique=True)
    code = models.CharField(max_length=20, unique=True)
    
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='division_created_by', blank=True, null=True) 
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='division_updated_by', blank=True, null=True) 
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
 
    def __str__(self):
        return self.name


class SubDivision(models.Model):
    division = models.ForeignKey(Division, on_delete=models.CASCADE, related_name="subdivisions")
    name     = models.CharField(max_length=120)
    code     = models.CharField(max_length=20)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='subdivision_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='subdivision_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
 
    def __str__(self):
        return f"{self.division.code} | {self.name}"
 
 
class Section(models.Model):
    subdivision = models.ForeignKey(SubDivision, on_delete=models.CASCADE, related_name="sections")
    name        = models.CharField(max_length=120)
    code        = models.CharField(max_length=20)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='section_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='section_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    class Meta:
        unique_together = [("subdivision", "code")]
        ordering = ["name"]
 
    def __str__(self):
        return f"{self.subdivision} | {self.name}"
 
 
class Building(models.Model): 
    STATUS_CHOICES = [
        ("Active", "Active"),
        ("Inactive", "Inactive"),
        ("Under Maintenance", "Under Maintenance"),
        ("Decommissioned", "Decommissioned"),
    ] 
 
    # Identity
    building_id  = models.CharField(max_length=50, unique=True, null=True, blank=True)
    name         = models.CharField(max_length=200, unique=True)
    surname      = models.CharField(max_length=50, blank=True) 
    year_of_establishment = models.DateTimeField(null=True, blank=True)

    # Hierarchy
    division     = models.ForeignKey(Division,    null=True, blank=True, on_delete=models.SET_NULL, related_name="buildings")
    subdivision  = models.ForeignKey(SubDivision, null=True, blank=True, on_delete=models.SET_NULL, related_name="buildings")
    section      = models.ForeignKey(Section,     null=True, blank=True, on_delete=models.SET_NULL, related_name="buildings")
 
    # Location
    address = models.TextField(blank=True)
    latitude         = models.DecimalField(max_digits=9, decimal_places=6, null=True, blank=True)
    longitude        = models.DecimalField(max_digits=9, decimal_places=6, null=True, blank=True)
 
    # Status
    status           = models.CharField(max_length=30, choices=STATUS_CHOICES, default="Active")

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='building_created_by', blank=True, null=True) 
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='building_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 

    def save(self, *args, **kwargs):
        if not self.building_id:
            base_id = slugify(self.name)[:30].upper()
            unique_suffix = uuid.uuid4().hex[:6].upper()
            self.building_id = f"{base_id}-{unique_suffix}"
        super().save(*args, **kwargs) 
 
 
    def __str__(self):
        label = self.surname or self.name
        return f"[{self.id}] — {label}"
 

class Technician(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name="technician_profile", null=True, blank=True)
    first_name = models.CharField(max_length=50, blank=True)
    last_name = models.CharField(max_length=50, blank=True)
    email = models.EmailField(blank=True) 
    phone_number = models.CharField(max_length=20, blank=True)
    about = models.TextField(blank=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='technician_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='technician_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    def __str__(self):
        return self.display_name

    @property
    def display_name(self):
        if self.user:
            return self.user.get_full_name() or self.user.username
        if self.first_name or self.last_name:
            return f"{self.first_name} {self.last_name}".strip()
        return "No User Assigned"

    @property
    def first_initial(self):
        name = self.display_name
        return name[0].upper() if name else "T"


class EquipmentType(models.Model):
    name = models.CharField(max_length=50, unique=True)
    description = models.TextField(blank=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_type_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_type_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    def __str__(self):
        return self.name
    
class EquipmentTypeData(models.Model):
    equipment_type = models.ForeignKey(EquipmentType, on_delete=models.CASCADE, related_name="data_fields")
    field_name     = models.CharField(max_length=100)
    field_type     = models.CharField(max_length=20, choices=[("text", "Text"), ("number", "Number"), ("date", "Date"), ("boolean", "Boolean")], default="text")
    is_required    = models.BooleanField(default=False)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_type_data_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_type_data_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 

 
    def __str__(self):
        return f"{self.equipment_type.name} - {self.field_name}"


class EquipmentComponentPreset(models.Model):
    equipment_type = models.ForeignKey(EquipmentType, on_delete=models.CASCADE, related_name="component_presets")
    name = models.CharField(max_length=100)  
    description = models.TextField(blank=True)
 
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_component_preset_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_component_preset_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    def __str__(self):
        return self.name 
 

class Equipment(models.Model):
    
    STATUS_CHOICES = [
        ("active",             "Active"),
        ("in_repair",          "In Repair"),
        ("under_maintenance",  "Under Maintenance"),
        ("inactive",           "Inactive"),
        ("decommissioned",     "Decommissioned"),
    ]

    equipment_id   = models.CharField(max_length=50, unique=True, blank=True) 
    equipment_type = models.ForeignKey(EquipmentType, on_delete=models.CASCADE, related_name="equipment")
    building       = models.ForeignKey(Building, on_delete=models.CASCADE, related_name="equipment")
    component_preset = models.ManyToManyField(EquipmentComponentPreset, related_name="equipment", blank=True)
    brand          = models.CharField(max_length=100, blank=True, help_text="e.g. Otis, Daikin")
    floor_location = models.CharField(max_length=100, blank=True, help_text="e.g. Level 4 Lobby")
    installment_date = models.DateTimeField(null=True, blank=True)

    model_number   = models.CharField(max_length=100, blank=True)
    serial_number  = models.CharField(max_length=100, blank=True)

    installation_year = models.PositiveSmallIntegerField(null=True, blank=True)
    qr_code        = models.CharField(max_length=100, blank=True, null=True, unique=True) 

    image          = models.ImageField(upload_to="equipment/", null=True, blank=True)
    notes          = models.TextField(blank=True)
 
    # Status
    status         = models.CharField(max_length=30, choices=STATUS_CHOICES, default="active")
 
    # Maintenance scheduling 
    maintenance_period_days = models.PositiveIntegerField(default=90,)
    last_maintenance_date   = models.DateField(null=True, blank=True)
    next_service_due        = models.DateField(null=True, blank=True, db_index=True,)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True) 
    deleted = models.BooleanField(default=False)


    def save(self, *args, **kwargs):
        if not self.qr_code:
            self.qr_code = f"EQ-{self.equipment_id or uuid.uuid4().hex[:8].upper()}"
        super().save(*args, **kwargs)
    
    def __str__(self):
        return f"[{self.equipment_id}] {self.equipment_type.name} at {self.building.name}"
 


class EquipmentComponents(models.Model):
    equipment = models.ForeignKey(Equipment, on_delete=models.CASCADE, related_name="components")
    name      = models.CharField(max_length=100)
    description = models.TextField(blank=True) 

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_component_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='equipment_component_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    def __str__(self):
        return f"{self.name} ({self.equipment.equipment_id})"


class Maintenance(models.Model):
    maintenance_serial = models.CharField(max_length=100, blank=True, unique=True, null= True)
    qr_code = models.CharField(max_length=100, blank=True, unique=True) 
    building = models.ForeignKey(Building, on_delete=models.SET_NULL, null=True, blank=True, related_name="maintenance_sessions")
    equipment = models.ManyToManyField(Equipment, related_name="maintenance_sessions", blank=True)
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False)

    def save(self, *args, **kwargs):
        if not self.maintenance_serial:
            self.maintenance_serial = f"MT-{uuid.uuid4().hex[:8].upper()}" 
        if not self.qr_code:
            self.qr_code = f"QR-MT-{uuid.uuid4().hex[:8].upper()}"
        super().save(*args, **kwargs)

    def __str__(self):
        return f"Maintenance {self.maintenance_serial}"


class MaintenanceRecord(models.Model):
    RECORD_TYPE_CHOICES = [
        ("routine", "Routine"),
        ("emergency", "Emergency"),
        ("repair", "Repair"),
        ("inspection", "Inspection"),
        ("upgrade", "Upgrade"),
    ]

    REVIEW_STATUS_CHOICES = [
        ("pending", "Pending Review"),
        ("approved", "Approved"),
        ("rejected", "Rejected"),
    ]

    MAINTAIANCE_STATUS_CHOICES = [
        ("pending", "Pending"),
        ("in_progress", "In Progress"),
        ("completed", "Completed"),
        ("skipped", "Skipped"),
        ("cancelled", "Cancelled"),
    ] 

    # Core relation
    equipment    = models.ForeignKey(Equipment, on_delete=models.PROTECT, related_name="maintenance_records")
    maintenance  = models.ForeignKey(Maintenance, on_delete=models.CASCADE, related_name="maintenance_records", null=True, blank=True) 
    ticket       = models.ForeignKey('backend.Ticket', on_delete=models.SET_NULL, null=True, blank=True, related_name="maintenance_records")
    
    maintainance_status = models.CharField(max_length=20, choices=MAINTAIANCE_STATUS_CHOICES, default="pending", db_index=True) 
    opt         = models.CharField(max_length=255, blank=True) 
    assigned_to   = models.ForeignKey(Technician, on_delete=models.SET_NULL, null=True, blank=True, related_name="assigned_maintenance_records") 

    # Record metadata
    record_type       = models.CharField(max_length=20, choices=RECORD_TYPE_CHOICES, default="routine")
    maintenance_date  = models.DateField(db_index=True)
    technician        = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="maintenance_records",)
    work_description  = models.TextField()
    parts_replaced    = models.TextField(blank=True)
    cost              = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, help_text="Cost in local currency.")
    duration_hours    = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
 
    # Approval flow
    review_status     = models.CharField(max_length=10, choices=REVIEW_STATUS_CHOICES, default="pending", db_index=True)
    reviewed_by       = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="reviewed_maintenance_records")
    reviewed_at       = models.DateTimeField(null=True, blank=True)
    review_notes      = models.TextField(blank=True, help_text="Reviewer comment (required on rejection).")
    

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_record_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_record_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)    
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 

    def save(self, *args, **kwargs):
        if not self.opt:
            self.opt = ''.join(random.choices(string.digits, k=6))
        super().save(*args, **kwargs) 

    def __str__(self):
        return f"{self.record_type.title()} maintenance for {self.equipment} on {self.maintenance_date}" 
 

class MaintenanceComponent(models.Model):
    maintenance_record = models.ForeignKey(MaintenanceRecord, on_delete=models.CASCADE, related_name="components")
    name               = models.CharField(max_length=100)
    description        = models.TextField(blank=True) 
    is_checked         = models.BooleanField(default=False)
    remark             = models.TextField(blank=True, null=True)
    suggestion         = models.TextField(blank=True, null=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_component_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_component_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    def __str__(self):
        return f"{self.name} ({self.maintenance_record})" 
 
 
class MaintenanceAttachment(models.Model):
    maintenance = models.ForeignKey(Maintenance, on_delete=models.CASCADE, related_name="attachments", null=True, blank=True)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name="maintenance_attachments")
    file        = models.FileField(upload_to="maintenance/attachments/%Y/%m/")
    caption     = models.CharField(max_length=200, blank=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_attachment_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='maintenance_attachment_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)    

    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False)

    def __str__(self):
        return f"Attachment for {self.maintenance} by {self.uploaded_by}" 
 


class Ticket(models.Model): 
    STATUS_CHOICES = [
        ("open", "Open"),
        ("assigned", "Assigned"),
        ("in_progress", "In Progress"),
        ("resolved", "Resolved"),
        ("closed", "Closed | Done"),
        ("rejected", "Rejected"),
    ]

    PRIORITY_CHOICES = [
        ("critical", "Critical"),
        ("high", "High"),
        ("medium", "Medium"),
        ("low", "Low"),
    ]

    ISSUE_TYPE_CHOICES = [
        ("breakdown", "Breakdown"),
        ("routine", "Routine Check"),
        ("emergency", "Emergency"),
        ("inspection", "Inspection"),
        ("warranty", "Warranty"),
        ("other", "Other"),
    ]

    user        = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name="reported_tickets") 
    equipment   = models.ForeignKey(Equipment, on_delete=models.PROTECT, related_name="tickets")
    building    = models.ForeignKey(Building, on_delete=models.PROTECT, related_name="tickets",  help_text="Denormalised from equipment.building for faster filtering.")
 
    status     = models.CharField(max_length=20, choices=STATUS_CHOICES, default="open", db_index=True)
    priority   = models.CharField(max_length=10, choices=PRIORITY_CHOICES, default="medium", db_index=True)
    issue_type = models.CharField(max_length=20, choices=ISSUE_TYPE_CHOICES, default="breakdown", db_index=True) 

    # Description
    title             = models.CharField(max_length=255)
    issue_description = models.TextField(blank=True)
 
    # Timestamps
    opened_at    = models.DateTimeField(default=timezone.now)
    assigned_at  = models.DateTimeField(null=True, blank=True)
    resolved_at  = models.DateTimeField(null=True, blank=True)
    closed_at    = models.DateTimeField(null=True, blank=True)
 
    # SLA | scheduling
    due_date     = models.DateTimeField(null=True, blank=True)
    estimated_hours = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    actual_hours    = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    notes        = models.TextField(blank=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='ticket_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='ticket_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False)


    def __str__(self):
        return f"{self.id}: {self.title}"
 
 
class TicketComment(models.Model):
    ticket  = models.ForeignKey(Ticket, on_delete=models.CASCADE, related_name="comments")
    author  = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name="ticket_comments")
    body    = models.TextField()
    is_internal = models.BooleanField(default=False,)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='ticket_comment_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='ticket_comment_updated_by', blank=True, null=True) 
    created_at  = models.DateTimeField(auto_now_add=True)
    updated_at  = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 
 
    class Meta:
        ordering  = ["created_at"]
 
    def __str__(self):
        return f"Comment on {self.ticket.id} by {self.author}"
 

class TicketActivityLog(models.Model):
    ACTION_TYPES = [
        ("created", "Created Ticket"),
        ("status_changed", "Status Changed"),
        ("assigned", "Technician Assigned"),
        ("comment_added", "Comment Added"),
        ("priority_changed", "Priority Changed"),
        ("resolved", "Marked Resolved"),
        ("closed", "Closed"),
    ]
 
    ticket      = models.ForeignKey(Ticket, on_delete=models.CASCADE, related_name="activity_log")
    actor       = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name="ticket_actions")
    action_type = models.CharField(max_length=30, choices=ACTION_TYPES)
    description = models.TextField()
    metadata    = models.JSONField(default=dict, blank=True,)
    created_at  = models.DateTimeField(default=timezone.now)


 
    def __str__(self):
        return f"{self.action_type} on {self.ticket.id} @ {self.created_at:%Y-%m-%d %H:%M}"
    

class AssignTechnician(models.Model):
    ticket = models.ForeignKey(Ticket, on_delete=models.CASCADE, related_name="assigned_technicians")
    technician = models.ForeignKey(Technician, on_delete=models.SET_NULL, null=True, blank=True, related_name="assigned_tickets") 
    assign_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="ticket_assignments") 
    assigned_at = models.DateTimeField(auto_now_add=True)
    is_active = models.BooleanField(default=True) 
    deleted = models.BooleanField(default=False)

    def __str__(self):
        return f"{self.technician}" 



class AssignActivities(models.Model):

    STATUS_CHOICES = [
        ("pending", "Pending"),
        ("accepted", "Accepted"),
        ("rejected", "Rejected"),
    ]

    COMPLATE_STATUS_CHOICES = [
        ("pending", "Pending"),
        ("completed", "Completed"),
        ("skipped", "Skipped"),
        ('overdue', "Overdue"),
        ("cancelled", "Cancelled"),
    ]

    technician = models.ForeignKey(Technician, on_delete=models.CASCADE, related_name="activities")
    ticket = models.ForeignKey(Ticket, on_delete=models.CASCADE, related_name="activities")
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="pending")
    completion_status = models.CharField(max_length=20, choices=COMPLATE_STATUS_CHOICES, default="pending")
    
    assigned_at = models.DateTimeField(auto_now_add=True)
    completed_at = models.DateTimeField(null=True, blank=True) 
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='assign_activity_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='assign_activity_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True) 



class ScheduledMaintenance(models.Model):
    RECURRENCE_CHOICES = [
        ("once", "Once"),
        ("daily", "Daily"),
        ("weekly", "Weekly"),
        ("monthly", "Monthly"),
        ("quarterly", "Quarterly"),
        ("biannual", "Bi-annual"),
        ("annual", "Annual"),
    ]

    SCHEDUL_STATUS_CHOICES = [
        ("pending", "Pending"),
        ("completed", "Completed"),
        ("skipped", "Skipped"),
        ("cancelled", "Cancelled"),
    ]
    equipment        = models.ForeignKey(Equipment, on_delete=models.CASCADE, related_name="scheduled_maintenances")
    title            = models.CharField(max_length=200)
    description      = models.TextField(blank=True)
    scheduled_date   = models.DateField(db_index=True)

    recurrence       = models.CharField(max_length=15, choices=RECURRENCE_CHOICES, default="once")
    assigned_to      = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="scheduled_maintenances") 
    status           = models.CharField(max_length=15, choices=SCHEDUL_STATUS_CHOICES, default="pending", db_index=True)
    completed_record = models.OneToOneField(MaintenanceRecord, on_delete=models.SET_NULL, null=True, blank=True, related_name="scheduled_maintenance", help_text="Filled in when this schedule is fulfilled.",)
    estimated_cost   = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)

    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='scheduled_maintenance_created_by', blank=True, null=True)
    updated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='scheduled_maintenance_updated_by', blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_active = models.BooleanField(default=True)
    deleted = models.BooleanField(default=False) 

 
    def __str__(self):
        return f"Scheduled: {self.title} for {self.equipment} on {self.scheduled_date}"
 

class CriticalAlert(models.Model):

    SEVERITY_CHOICES = [
        ("critical", "Critical"),
        ("high", "High"),
        ("medium", "Medium"),
        ("low", "Low"),
    ]

    ALERTYPES_CHOICES = [
        ("warranty", "Warranty Expiry"),
        ("maint_overdue", "Maintenance Overdue"),
        ("eq_fault", "Equipment Fault"),
        ("repair", "Repair Pending"),
        ("inspection", "Inspection Due"),
        ("sla_breach", "Ticket SLA Breach"),
        ("custom", "Custom"),
    ]

    equipment       = models.ForeignKey(Equipment, on_delete=models.CASCADE, related_name="alerts", null=True, blank=True,)
    building        = models.ForeignKey(Building, on_delete=models.CASCADE, related_name="alerts", null=True, blank=True,) 
    alert_type      = models.CharField(max_length=20, choices=ALERTYPES_CHOICES)
    severity        = models.CharField(max_length=10, choices=SEVERITY_CHOICES, default="high", db_index=True)
    title           = models.CharField(max_length=255)
    description     = models.TextField(blank=True)
    is_acknowledged = models.BooleanField(default=False, db_index=True)

    acknowledged_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="acknowledged_alerts")
    acknowledged_at = models.DateTimeField(null=True, blank=True)
    expires_at      = models.DateTimeField(null=True, blank=True)
    metadata        = models.JSONField(default=dict, blank=True)
 
 
    def __str__(self):
        return f"[{self.severity.upper()}] {self.title}"

