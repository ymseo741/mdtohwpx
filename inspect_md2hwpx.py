import md2hwpx
import inspect

print("--- md2hwpx module members ---")
for name, obj in inspect.getmembers(md2hwpx):
    if not name.startswith('_'):
        print(f"{name}: {type(obj)}")

print("\n--- help(md2hwpx) ---")
help(md2hwpx)
