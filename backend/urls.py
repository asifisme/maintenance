from django.urls import path, include 
from backend import views 

app_name = 'backend' 

urlpatterns = [
    path("api/menu-search/", views.search_backend_menus, name="search_backend_menus"),

    # Dashboards
    # Image optimization
    path("image/<str:unique_key>/", views.serve_optimized_image, name="serve_optimized_image"),

    path('login/', views.backend_login, name='backend_login'),
    path('logout/', views.backend_logout, name='backend_logout'),

    path('<str:menu_slug>-menu/', views.menu_wise_dashboard, name='menu_wise_dashboard'),

    # User Management
    path('user/', views.UserListView.as_view(), name='user_list'),
    path('user/add/', views.user_add, name='user_add'),
    path('user/update/<str:data_id>/', views.user_update, name='user_update'),
    path('user/password/reset/<str:data_id>/', views.reset_password, name='reset_password'),
    path('user/permission/<int:user_id>/', views.user_permission, name='user_permission'),

    # Division
    path('division/', views.division_list, name='division_list'),
    path('division/add/', views.division_add, name='division_add'),
    path('division/update/<int:data_id>/', views.division_update, name='division_update'),
    path('division/delete/<int:data_id>/', views.division_delete, name='division_delete'),

    # SubDivision
    path('subdivision/', views.subdivision_list, name='subdivision_list'),
    path('subdivision/add/', views.subdivision_add, name='subdivision_add'),
   
]
