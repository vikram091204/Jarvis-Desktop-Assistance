from difflib import SequenceMatcher

# Test the command matching logic
query = "open browser"
cmd_variants = ["open"]

print(f"Query: '{query}'")
print(f"Command variants: {cmd_variants}")
print()

for variant in cmd_variants:
    print(f"Testing variant: '{variant}'")
    print(f"  variant in query: {variant in query}")
    print(f"  query in variant: {query in variant}")
    
    if variant in query or query in variant:
        ratio = SequenceMatcher(None, query, variant).ratio()
        print(f"  Match found! Ratio: {ratio}")
    else:
        ratio = SequenceMatcher(None, query, variant).ratio()
        print(f"  No substring match. Fuzzy ratio: {ratio}")
    print()

# Test what happens after stripping "open "
app = query.lower()
for prefix in ("open ", "please open ", "could you open "):
    if app.startswith(prefix):
        app = app[len(prefix):].strip()
        print(f"After stripping '{prefix}': '{app}'")

# Test platform matching
platforms = {'browser': 'https://www.google.com'}
print(f"\nTesting platform matching for '{app}':")
for platform, url in platforms.items():
    print(f"  Platform: '{platform}'")
    print(f"  platform in app: {platform in app}")
    print(f"  app in platform: {app in platform}")
    ratio = SequenceMatcher(None, app, platform).ratio()
    print(f"  Fuzzy ratio: {ratio}")
