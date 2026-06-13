from django import forms
from django.contrib.auth.models import User
from backend.models import (
    WebImages,
    Maintenance,
    Equipment,
    Building,
)

TAILWIND_TEXT = (
    "h-11 w-full mt-1 rounded-lg bg-[var(--bg-color)] px-3 text-sm "
    "ring-1 ring-[var(--border-color)] outline-none "
    "focus:ring-[var(--primary-color)]/40"
)

TAILWIND_TEXTAREA = (
    "h-11 w-full mt-1 pt-2 rounded-lg bg-[var(--bg-color)] px-3 text-sm "
    "ring-1 ring-[var(--border-color)] outline-none "
    "focus:ring-[var(--primary-color)]/40"
)

TAILWIND_SELECT = (
    "h-11 w-full mt-1 appearance-none rounded-lg bg-[var(--bg-color)] px-3 pr-10 text-sm "
    "ring-1 ring-[var(--border-color)] outline-none "
    "focus:ring-[var(--primary-color)]/40"
)

GENDER_CHOICES = (
    ("male", "Male"),
    ("female", "Female"),
    ("other", "Other"),
)




class CustomUserLoginForm(forms.Form):
    """
    Custom form for user login.
    This form is used to authenticate users in the backend.
    """
    username = forms.CharField(
        max_length=150, required=True,
        widget=forms.TextInput(attrs={'placeholder': 'Enter Your Username', 'class': 'h-14 w-full rounded-md border border-slate-200 bg-[#A5B4FC26] pl-9 pr-3 text-sm outline-none transition focus:border-brand-600 focus:bg-white'})
    )
    password = forms.CharField(
        max_length=128, required=True,
        widget=forms.PasswordInput(attrs={'placeholder': 'Enter Your Password', 'class': 'h-14 w-full rounded-md border border-slate-200 bg-[#A5B4FC26] pl-9 pr-10 text-sm outline-none transition focus:border-brand-600 focus:bg-white'})
    )

    def clean_username(self):
        username = self.cleaned_data.get('username')
        return username.strip() if username else username

    def clean_password(self):
        password = self.cleaned_data.get('password')
        return password




class UserCreateForm(forms.ModelForm):
    
    first_name = forms.CharField(
        label="Full Name",
        max_length=100,
        widget=forms.TextInput(attrs={
            "placeholder": "Enter full name",
            "class": TAILWIND_TEXT,
            "id": "first_name",
        })
    )
    email = forms.EmailField(
        label="Email",
        widget=forms.EmailInput(attrs={
            "placeholder": "Enter email",
            "class": TAILWIND_TEXT,
            "id": "email",
        })
    )

    last_name = forms.CharField(required=False, widget=forms.TextInput(attrs={"class": TAILWIND_TEXT}))
    date_of_birth = forms.DateField(required=False, widget=forms.DateInput(attrs={
        "type": "date",
        "class": TAILWIND_TEXT,
    }))

    phone = forms.CharField(
        label="Mobile No.",
        max_length=30,
        widget=forms.TextInput(attrs={
            "placeholder": "Enter phone number",
            "class": TAILWIND_TEXT,
            "id": "phone",
        })
    )
    gender = forms.ChoiceField(
        choices=GENDER_CHOICES,
        widget=forms.Select(attrs={
            "class": f'{TAILWIND_SELECT} select2-items',
            "id": "gender",
        })
    )
    profile_image = forms.ImageField(
        label="Profile Image",
        required=False,
        widget=forms.ClearableFileInput(attrs={
            "accept": "image/*",
            "class": "sr-only",
            "id": "profile_image",
        })
    )

    class Meta:
        model = User
        fields = ["phone", "gender", "profile_image"]

    def clean_email(self):
        email = self.cleaned_data.get("email")
        user_id = self.instance.user.pk if getattr(self.instance, "user", None) else None
        qs = User.objects.exclude(pk=user_id) if user_id else User.objects.all()
        if qs.filter(email=email).exists():
            raise forms.ValidationError("This email is already taken.")
        return email

    def save(self, commit=True):
        admin_user = super().save(commit=False)
        user = getattr(self.instance, "user", None)
        if not user:
            user = User()

        user.first_name = self.cleaned_data.get("first_name", "")
        user.last_name = self.cleaned_data.get("last_name", "")
        user.email = self.cleaned_data.get("email")
        user.username = user.email

        if not user.pk:
            user.set_password("12345678")

        if commit:
            user.save()
            admin_user.user = user
            admin_user.save()

        return admin_user


class MaintenanceForm(forms.ModelForm):
    """Form for creating and editing Maintenance sessions"""
    
    building = forms.ModelChoiceField(
        queryset=Building.objects.filter(deleted=False, is_active=True),
        required=True,
        widget=forms.Select(attrs={
            'class': f'{TAILWIND_SELECT} select2-items',
            'id': 'building_select',
        }),
        label="Building"
    )
    
    equipment = forms.ModelMultipleChoiceField(
        queryset=Equipment.objects.none(),
        required=False,
        widget=forms.CheckboxSelectMultiple(attrs={
            'class': 'equipment-checkbox',
            'id': 'equipment_select',
        }),
        label="Equipment"
    )
    
    maintenance_serial = forms.CharField(
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={
            'placeholder': 'Leave empty to auto-generate (e.g. MT-XXXXXX)',
            'class': TAILWIND_TEXT,
        }),
        label="Maintenance Serial"
    )
    
    qr_code = forms.CharField(
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={
            'placeholder': 'Leave empty to auto-generate QR code',
            'class': TAILWIND_TEXT,
        }),
        label="QR Code Value"
    )
    
    class Meta:
        model = Maintenance
        fields = ['building', 'equipment', 'maintenance_serial', 'qr_code']
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # If editing, load equipment for the selected building
        if self.instance and self.instance.building_id:
            self.fields['equipment'].queryset = Equipment.objects.filter(
                building=self.instance.building,
                deleted=False,
                is_active=True
            )
