Custom organization-specific extensions live here.

Purpose:
- Keep local company features outside the upstream `Extensions` and `Xaml` folders.
- Make future updates easier by preserving custom code in `Custom`.

Current customizations:
- Menu search box
- Persistent menu favorites

Structure:
- `Custom\Extensions\*.psm1`: custom PowerShell modules, auto-imported after the standard modules
- `Custom\Xaml\...`: XAML overrides; if a file exists here with the same relative path as the upstream file, it is loaded instead

Current modules:
- `Custom\Extensions\MenuSearch.Custom.psm1`
- `Custom\Extensions\MenuFavorites.Custom.psm1`

Update workflow:
1. Update the upstream project files as usual.
2. Keep the `Custom` folder.
3. Start the app and verify that the custom hooks still run.
4. Only if upstream changed the relevant core hook points, adjust the small hook integration in `Core.psm1`.
