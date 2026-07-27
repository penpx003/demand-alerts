"""Weekly Demand Planning digest service."""

# Use the OS (Windows) certificate store so corporate TLS-interception root CAs
# are trusted. Must run before any HTTPS client builds its SSL context, which is
# why it lives here rather than in the app module — importing anything from this
# package injects it first. Without it, Supabase calls fail on the corporate
# network with "self-signed certificate in certificate chain".
try:
    import truststore

    truststore.inject_into_ssl()
except Exception:  # truststore optional; ignore if unavailable
    pass
