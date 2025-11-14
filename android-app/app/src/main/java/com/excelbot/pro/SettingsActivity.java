package com.excelbot.pro;

import android.app.Activity;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

public class SettingsActivity extends Activity {

    private EditText serverUrlInput;
    private Button saveButton;
    private Button resetButton;
    private SharedPreferences preferences;
    private static final String PREFS_NAME = "ExcelBotPrefs";
    private static final String SERVER_URL_KEY = "server_url";
    private static final String DEFAULT_SERVER_URL = "https://huggingface.co/spaces/YOUR_USERNAME/excelbot-pro";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_settings);

        preferences = getSharedPreferences(PREFS_NAME, MODE_PRIVATE);
        
        serverUrlInput = findViewById(R.id.serverUrlInput);
        saveButton = findViewById(R.id.saveButton);
        resetButton = findViewById(R.id.resetButton);

        // Load current server URL
        String currentUrl = preferences.getString(SERVER_URL_KEY, DEFAULT_SERVER_URL);
        serverUrlInput.setText(currentUrl);

        saveButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                saveSettings();
            }
        });

        resetButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                resetSettings();
            }
        });
    }

    private void saveSettings() {
        String serverUrl = serverUrlInput.getText().toString().trim();
        
        if (serverUrl.isEmpty()) {
            Toast.makeText(this, "Please enter a valid URL", Toast.LENGTH_SHORT).show();
            return;
        }

        SharedPreferences.Editor editor = preferences.edit();
        editor.putString(SERVER_URL_KEY, serverUrl);
        editor.apply();

        Toast.makeText(this, "Settings saved! Restart the app to apply changes.", Toast.LENGTH_LONG).show();
        finish();
    }

    private void resetSettings() {
        serverUrlInput.setText(DEFAULT_SERVER_URL);
        SharedPreferences.Editor editor = preferences.edit();
        editor.putString(SERVER_URL_KEY, DEFAULT_SERVER_URL);
        editor.apply();
        Toast.makeText(this, "Settings reset to default", Toast.LENGTH_SHORT).show();
    }
}
