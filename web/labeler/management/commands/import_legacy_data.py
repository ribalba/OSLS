from django.core.management.base import BaseCommand

from labeler.services import seed_from_legacy_files


class Command(BaseCommand):
    help = "Import legacy JSON config and print_log.jsonl into the Django database."

    def add_arguments(self, parser):
        parser.add_argument("--force", action="store_true", help="Replace existing imported data.")

    def handle(self, *args, **options):
        seed_from_legacy_files(force=options["force"])
        self.stdout.write(self.style.SUCCESS("Imported OSLS legacy data."))
