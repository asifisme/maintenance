from django.urls import path, include 
from backend import views 

app_name = 'backend' 

urlpatterns = [
    path("api/menu-search/", views.search_backend_menus, name="search_backend_menus"),

    # Dashboards
    path('', views.dash_board, name='dash_board'),
    path('dashboard/', views.dash_board, name='backend_dashboard'),
    path('dashboard/menu/<str:menu_slug>/', views.menu_wise_dashboard, name='menu_wise_dashboard'),

    # Image optimization
    path("image/<str:unique_key>/", views.serve_optimized_image, name="serve_optimized_image"),

    path('login/', views.backend_login, name='backend_login'),
    path('logout/', views.backend_logout, name='backend_logout'),


    # User Management
    path('user/', views.UserListView.as_view(), name='user_list'),
    path('user/add/', views.user_add, name='user_add'),
    path('user/update/<str:data_id>/', views.user_update, name='user_update'),
    path('user/password/reset/<str:data_id>/', views.reset_password, name='reset_password'),
    path('user/permission/<int:user_id>/', views.user_permission, name='user_permission'),

    # Technician Management
    path('technician/', views.technician_list, name='technician_list'),
    path('technician/add/', views.technician_add, name='technician_add'),
    path('technician/update/<int:data_id>/', views.technician_update, name='technician_update'),
    path('technician/delete/<int:data_id>/', views.technician_delete, name='technician_delete'),

    # Division
    path('division/', views.division_list, name='division_list'),
    path('division/add/', views.division_add, name='division_add'),
    path('division/update/<int:data_id>/', views.division_update, name='division_update'),
    path('division/status/<int:data_id>/', views.division_status, name='division_status'),

    # SubDivision
    path('subdivision/', views.subdivision_list, name='subdivision_list'),
    path('subdivision/add/', views.subdivision_add, name='subdivision_add'),
    path('subdivision/update/<int:data_id>/', views.subdivision_update, name='subdivision_update'),
    path('subdivision/status/<int:data_id>/', views.subdivision_status, name='subdivision_status'),

    # Section
    path('section/', views.section_list, name='section_list'),
    path('section/add/', views.section_add, name='section_add'),
    path('section/update/<int:data_id>/', views.section_update, name='section_update'),
    path('section/status/<int:data_id>/', views.section_status, name='section_status'),

    # Building
    path('building/', views.building_list, name='building_list'),
    path('building/add/', views.building_add, name='building_add'),
    path('building/update/<int:data_id>/', views.building_update, name='building_update'),
    path('building/details/<int:data_id>/', views.building_details, name='building_details'),
    path('building/status/<int:data_id>/', views.building_status, name='building_status'),

    # Equipment Type
    path('equipment-type/', views.equipment_type_list, name='equipment_type_list'),
    path('equipment-type/add/', views.equipment_type_add, name='equipment_type_add'),
    path('equipment-type/update/<int:data_id>/', views.equipment_type_update, name='equipment_type_update'),
    path('equipment-type/status/<int:data_id>/', views.equipment_type_status, name='equipment_type_status'),

    # Equipment Type Data
    path('equipment-type-data/', views.equipment_type_data_list, name='equipment_type_data_list'),
    path('equipment-type-data/add/', views.equipment_type_data_add, name='equipment_type_data_add'),
    path('equipment-type-data/update/<int:data_id>/', views.equipment_type_data_update, name='equipment_type_data_update'),
    path('equipment-type-data/status/<int:data_id>/', views.equipment_type_data_status, name='equipment_type_data_status'),

    # Equipment
    path('equipment/', views.equipment_list, name='equipment_list'),
    path('equipment/add/', views.equipment_add, name='equipment_add'),
    path('equipment/detail/<int:data_id>/', views.equipment_detail, name='equipment_detail'),
    path('equipment/qr/<int:data_id>/', views.qr_equipement_views, name='equipment_qr'),
    path('equipment/update/<int:data_id>/', views.equipment_update, name='equipment_update'),
    path('equipment/status/<int:data_id>/', views.equipment_status, name='equipment_status'),
    path('equipment/history/', views.equipment_history, name='equipment_history'),

    # Maintenance Record Review (approve / reject)
    path('maintenance/',views.maintenance_list,name='maintenance'),
    path('maintenance/add/',views.maintenance_add,name='maintenance_add'),
    path('maintenance/detail/<int:data_id>/',views.maintenance_detail,name='maintenance_detail'),
    path('maintenance/update/<int:data_id>/',views.maintenance_update,name='maintenance_update'),
    path('maintenance/status/<int:data_id>/',views.maintenance_status,name='maintenance_status'),
    path('maintenance/api/get-equipment/',views.get_building_equipment,name='get_building_equipment'),
    path('maintenance/api/get-equipment-components/', views.get_equipment_components, name='get_equipment_components'),
    path('maintenance/scan/<str:qr_code>/', views.maintanaince_scan, name='maintanaince_scan'),
    path('maintenance-record/<int:record_id>/review/', views.maintenance_record_review, name='maintenance_record_review'),

    # Ticket
    path('ticket/', views.ticket_list, name='ticket_list'),
    path('ticket/add/', views.ticket_add, name='ticket_add'),
    path('ticket/detail/<int:data_id>/', views.ticket_detail, name='ticket_detail'),
    path('ticket/update/<int:data_id>/', views.ticket_update, name='ticket_update'),
    path('ticket/status/<int:data_id>/', views.ticket_status, name='ticket_status'),
]
