import os
from django.core.management import call_command
from django.core.management.base import BaseCommand
from django.conf import settings


class Command(BaseCommand):
    help = 'Loads all baseline and project data fixtures in correct dependency order.'

    FIXTURE_FILES = [
        'fixtures/fixtures.json',
        'fixtures/data/divisions.json',
        'fixtures/data/subdivisions.json',
        'fixtures/data/sections.json',
        'fixtures/data/buildings.json',
        'fixtures/data/equipment_types.json',
        'fixtures/data/equipment_type_data.json',
        'fixtures/data/equipment_component_presets.json',
        'fixtures/data/equipments.json',
        'fixtures/data/technicians.json',
        'fixtures/data/tickets.json',
        'fixtures/data/ticket_comments.json',
        'fixtures/data/maintenances.json',
        'fixtures/data/maintenance_records.json',
        'fixtures/data/maintenance_components.json',
        'fixtures/data/scheduled_maintenances.json',
        'fixtures/data/critical_alerts.json',
    ]

    def handle(self, *args, **options):
        base_dir = str(settings.BASE_DIR)
        valid_fixtures = []

        for rel_path in self.FIXTURE_FILES:
            full_path = os.path.join(base_dir, rel_path)
            if os.path.exists(full_path):
                valid_fixtures.append(full_path)
            else:
                self.stdout.write(self.style.WARNING(f'Fixture file not found: {rel_path}'))

        if not valid_fixtures:
            self.stdout.write(self.style.ERROR('No fixture files found to load.'))
            return

        self.stdout.write(self.style.NOTICE(f'Loading {len(valid_fixtures)} fixture files in dependency order...'))
        try:
            call_command('loaddata', *valid_fixtures)
            self.stdout.write(self.style.SUCCESS(f'Successfully loaded all {len(valid_fixtures)} fixtures.'))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error loading fixtures: {e}'))
