# ExcelBot Pro - Android App

This is the Android version of ExcelBot Pro, a VBA macro generator for Excel.

## Architecture

This Android app uses a **WebView-based approach**:
- The app displays the Gradio web interface inside an Android WebView
- The Python backend runs on a cloud server (Hugging Face Spaces)
- The app simply connects to the cloud URL

## Project Structure

```
android-app/
├── app/
│   ├── src/
│   │   └── main/
│   │       ├── java/com/excelbot/pro/
│   │       │   ├── MainActivity.java          # Main app screen
│   │       │   └── SettingsActivity.java      # Settings screen
│   │       ├── res/
│   │       │   ├── layout/
│   │       │   │   ├── activity_main.xml      # Main layout
│   │       │   │   └── activity_settings.xml  # Settings layout
│   │       │   ├── values/
│   │       │   │   ├── strings.xml            # String resources
│   │       │   │   └── styles.xml             # App theme
│   │       │   └── menu/
│   │       │       └── main_menu.xml          # Menu items
│   │       └── AndroidManifest.xml            # App configuration
│   └── build.gradle                            # App-level build config
├── build.gradle                                # Project-level build config
├── settings.gradle                             # Gradle settings
└── gradle.properties                           # Gradle properties
```

## Building the App

### Prerequisites
- Android Studio (latest version)
- JDK 8 or higher
- Android SDK (API 21-33)

### Steps
1. Open this folder in Android Studio
2. Wait for Gradle sync
3. Update `MainActivity.java` with your server URL
4. Build → Generate Signed Bundle / APK
5. Select APK and build

## Key Features

### MainActivity.java
- WebView configuration with JavaScript enabled
- Network connectivity checking
- Progress bar for loading states
- Error handling with user-friendly dialogs
- Back button navigation support

### SettingsActivity.java
- Configure custom server URL
- Save settings using SharedPreferences
- Reset to default URL option

### Permissions
- `INTERNET` - Required to load web content
- `ACCESS_NETWORK_STATE` - Check internet connectivity
- `READ/WRITE_EXTERNAL_STORAGE` - For file operations (future use)

## Configuration

### Default Server URL
Located in `MainActivity.java`:
```java
private static final String DEFAULT_SERVER_URL = "https://huggingface.co/spaces/YOUR_USERNAME/excelbot-pro";
```

Change `YOUR_USERNAME` to your Hugging Face username before building.

## Customization

### Change App Name
Edit `app/src/main/res/values/strings.xml`:
```xml
<string name="app_name">Your App Name</string>
```

### Change App Theme Colors
Edit `app/src/main/res/values/styles.xml`:
```xml
<item name="android:colorPrimary">#YourColor</item>
```

### Add App Icon
Replace files in:
- `app/src/main/res/mipmap-hdpi/ic_launcher.png`
- `app/src/main/res/mipmap-mdpi/ic_launcher.png`
- etc.

## Testing

### Debug Build
1. Connect Android device via USB
2. Enable USB debugging on device
3. Run → Run 'app' in Android Studio

### Release Build
1. Build → Generate Signed Bundle / APK
2. Create/use keystore
3. Build release APK
4. Test on device before distribution

## Distribution

### Option 1: Direct APK Distribution
- Share the APK file directly
- Users must enable "Install from Unknown Sources"

### Option 2: Google Play Store
- Requires Google Play Developer account ($25 one-time fee)
- Follow Play Store submission guidelines
- App review process takes 1-7 days

## Troubleshooting

### Gradle Sync Fails
- Check internet connection
- Update Android Studio
- Invalidate Caches → Restart

### App Crashes on Launch
- Check AndroidManifest.xml syntax
- Verify permissions are declared
- Check server URL is valid

### WebView Shows Blank Page
- Verify internet connection
- Check server URL is accessible
- Ensure JavaScript is enabled in WebView

## License

MIT License - Same as main ExcelBot Pro project

## Support

For detailed user instructions, see `ANDROID_USER_MANUAL.md` in the root folder.
