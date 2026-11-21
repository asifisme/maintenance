from django import template
from django.utils.safestring import mark_safe

register = template.Library()


@register.simple_tag(takes_context=True)
def get_page_object(context, explicit_page=None):
    """
    Safely retrieves the active Page object from context without triggering VariableDoesNotExist.
    Handles empty Page objects (where bool(page_obj) == False).
    """
    if explicit_page is not None and explicit_page != '':
        return explicit_page
    page_obj = context.get('page_obj')
    if page_obj is not None and page_obj != '':
        return page_obj
    data_list = context.get('data_list')
    if data_list is not None and data_list != '':
        return data_list
    return None


@register.simple_tag(takes_context=True)
def url_replace(context, **kwargs):
    """
    Preserve existing GET query parameters while updating/adding/removing specific keys.
    Usage: {% url_replace page=p %} or {% url_replace per_page=20 page=1 %}
    """
    request = context.get('request')
    if not request:
        query_items = []
        for k, v in kwargs.items():
            if v is not None and v != '':
                query_items.append(f"{k}={v}")
        return mark_safe(('?' + '&'.join(query_items)) if query_items else '?')

    query = request.GET.copy()
    for k, v in kwargs.items():
        if v is not None and v != '':
            query[k] = str(v)
        else:
            query.pop(k, None)

    qs = query.urlencode()
    return mark_safe(f"?{qs}" if qs else '?')


@register.filter
def pagination_window(page_obj, max_pages=5):
    """
    Returns a window of page numbers centered around current page (default 5 numbers),
    matching Figma inspect Frame 2147207612.
    """
    if not page_obj or not hasattr(page_obj, 'paginator'):
        return []
    num_pages = page_obj.paginator.num_pages
    current = page_obj.number
    if num_pages <= max_pages:
        return list(range(1, num_pages + 1))

    start = max(1, current - 2)
    end = min(num_pages, current + 2)

    if end - start < 4:
        if start == 1:
            end = min(num_pages, start + 4)
        elif end == num_pages:
            start = max(1, end - 4)

    return list(range(start, end + 1))


@register.filter
def page_item_style(page_num, current_page):
    """
    Returns 'active', 'adjacent', or 'outer' for exact styling in Figma:
    - active: black pill (#2B2823) with white text
    - adjacent: white pill with border and subtle drop shadow (#FFFFFF)
    - outer: subtle text without white background, elevated on hover
    """
    try:
        p = int(page_num)
        c = int(current_page)
        if p == c:
            return 'active'
        elif abs(p - c) == 1:
            return 'adjacent'
        return 'outer'
    except (ValueError, TypeError):
        return 'outer'


@register.simple_tag(takes_context=True)
def get_current_per_page(context):
    """
    Returns current per_page setting (default 10).
    """
    request = context.get('request')
    if request:
        val = request.GET.get('per_page')
        if val and str(val).isdigit() and int(val) > 0:
            return int(val)
    per_page = context.get('per_page')
    if per_page and str(per_page).isdigit() and int(per_page) > 0:
        return int(per_page)
    return 10
