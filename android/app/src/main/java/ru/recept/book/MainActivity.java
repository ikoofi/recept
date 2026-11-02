package ru.recept.book;

import android.app.Activity;
import android.Manifest;
import android.app.AlarmManager;
import android.app.AlertDialog;
import android.app.PendingIntent;
import android.content.ActivityNotFoundException;
import android.content.ClipData;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Insets;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.provider.MediaStore;
import android.view.Gravity;
import android.view.WindowInsets;
import android.view.WindowManager;
import android.webkit.CookieManager;
import android.webkit.JavascriptInterface;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.FrameLayout;
import android.widget.TextView;
import android.widget.Toast;

import androidx.core.content.FileProvider;

import com.chaquo.python.Python;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Окно приложения: запускает сайт на Python и показывает его в WebView. */
public class MainActivity extends Activity {
    private static final int REQUEST_PICK_FILE = 1;
    private static final int REQUEST_SAVE_FILE = 2;
    private static final String FILE_AUTHORITY = "ru.recept.book.files";

    /** Порт живёт вместе с процессом: при пересоздании окна сервер запускать не нужно. */
    private static int serverPort = 0;
    /** Секрет для сервера: без него запросы других приложений телефона отклоняются. */
    private static String accessToken = "";
    /** По нему TimerReceiver понимает, что сигнал подаст сама страница. */
    static volatile boolean isVisible = false;

    private FrameLayout root;
    private WebView webView;
    private ValueCallback<Uri[]> pendingFilePick;
    private String pendingDownloadUrl;
    private Uri pendingCameraPhoto;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        TextView loading = new TextView(this);
        loading.setText(R.string.app_name);
        loading.setTextSize(22);
        loading.setGravity(Gravity.CENTER);
        root = new FrameLayout(this);
        keepClearOfSystemBars(root);
        root.addView(loading);
        setContentView(root);

        // После пересоздания окна файл из исходного запроса уже импортирован.
        Uri incoming = savedInstanceState == null ? incomingFile(getIntent()) : null;
        new Thread(() -> {
            try {
                int port = startServer();
                boolean hasFile = incoming != null && copyIncoming(incoming);
                runOnUiThread(() -> showBook(port, hasFile));
            } catch (Exception e) {
                runOnUiThread(() -> loading.setText("Книга рецептов не запустилась:\n" + e));
            }
        }).start();
    }

    /**
     * На Android 15 окно рисуется под шторкой и панелью навигации, а клавиатура больше не
     * сжимает его сама. Отступаем на их размер. Отступ у самого WebView не сработает, он
     * игнорирует padding, поэтому он задаётся родителю.
     */
    private void keepClearOfSystemBars(FrameLayout view) {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.R) {
            return; // на старых версиях система сама не пускает окно под панели
        }
        view.setOnApplyWindowInsetsListener((v, insets) -> {
            Insets bars = insets.getInsets(WindowInsets.Type.systemBars()
                    | WindowInsets.Type.displayCutout()
                    | WindowInsets.Type.ime());
            v.setPadding(bars.left, bars.top, bars.right, bars.bottom);
            return WindowInsets.CONSUMED;
        });
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        setIntent(intent);
        Uri incoming = incomingFile(intent);
        if (incoming == null || webView == null) {
            return;
        }
        new Thread(() -> {
            if (copyIncoming(incoming)) {
                runOnUiThread(() -> webView.loadUrl(baseUrl(serverPort) + "import/incoming"));
            }
        }).start();
    }

    private static String baseUrl(int port) {
        return "http://127.0.0.1:" + port + "/";
    }

    /** Файл, которым с нами поделились или который открыли в книге. */
    @SuppressWarnings("deprecation")
    private static Uri incomingFile(Intent intent) {
        if (intent == null) {
            return null;
        }
        if (Intent.ACTION_VIEW.equals(intent.getAction())) {
            return intent.getData();
        }
        if (Intent.ACTION_SEND.equals(intent.getAction())) {
            return intent.getParcelableExtra(Intent.EXTRA_STREAM);
        }
        return null;
    }

    /** Кладёт файл туда, откуда его заберёт страница /import/incoming. */
    private boolean copyIncoming(Uri source) {
        File target = new File(new File(new File(getFilesDir(), "data"), "imports"), "incoming.json");
        try (InputStream in = getContentResolver().openInputStream(source);
             OutputStream out = new FileOutputStream(target)) {
            copy(in, out);
            return true;
        } catch (IOException | SecurityException e) {
            runOnUiThread(() -> Toast.makeText(this,
                    "Не удалось открыть файл: " + e.getMessage(), Toast.LENGTH_LONG).show());
            return false;
        }
    }

    private static void copy(InputStream in, OutputStream out) throws IOException {
        byte[] buffer = new byte[8192];
        int read;
        while ((read = in.read(buffer)) != -1) {
            out.write(buffer, 0, read);
        }
    }

    private synchronized int startServer() throws Exception {
        if (serverPort == 0) {
            File webapp = new File(getFilesDir(), "webapp");
            extractWebapp(webapp);
            File data = new File(getFilesDir(), "data");
            String[] started = Python.getInstance().getModule("android_main")
                    .callAttr("start", webapp.getPath(), data.getPath()).toString().split(" ");
            accessToken = started[1];
            serverPort = Integer.parseInt(started[0]);
        }
        return serverPort;
    }

    /** Код сайта лежит в assets, а Flask нужны обычные файлы. Распаковываем после каждого обновления. */
    private void extractWebapp(File target) throws Exception {
        String installed = String.valueOf(
                getPackageManager().getPackageInfo(getPackageName(), 0).lastUpdateTime);
        File marker = new File(target, ".installed");
        if (marker.exists()
                && installed.equals(new String(Files.readAllBytes(marker.toPath()), StandardCharsets.UTF_8))) {
            return;
        }
        deleteRecursively(target);
        copyAssets("webapp", target);
        Files.write(marker.toPath(), installed.getBytes(StandardCharsets.UTF_8));
    }

    private void copyAssets(String assetPath, File target) throws IOException {
        String[] children = getAssets().list(assetPath);
        if (children == null || children.length == 0) {
            try (InputStream in = getAssets().open(assetPath)) {
                Files.copy(in, target.toPath());
            }
            return;
        }
        target.mkdirs();
        for (String child : children) {
            copyAssets(assetPath + "/" + child, new File(target, child));
        }
    }

    private void deleteRecursively(File file) {
        File[] children = file.listFiles();
        if (children != null) {
            for (File child : children) {
                deleteRecursively(child);
            }
        }
        file.delete();
    }

    private void showBook(int port, boolean importFile) {
        webView = new WebView(this);
        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        CookieManager.getInstance().setAcceptCookie(true);
        webView.addJavascriptInterface(new AppBridge(), "AndroidApp");
        webView.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                if ("127.0.0.1".equals(request.getUrl().getHost())) {
                    return false;
                }
                // Ссылка на источник рецепта открывается в браузере телефона.
                try {
                    startActivity(new Intent(Intent.ACTION_VIEW, request.getUrl()));
                } catch (ActivityNotFoundException e) {
                    Toast.makeText(MainActivity.this, "Нет браузера", Toast.LENGTH_LONG).show();
                }
                return true;
            }

            @Override
            public void doUpdateVisitedHistory(WebView view, String url, boolean isReload) {
                updateKeepScreenOn(url);
            }
        });
        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView view, ValueCallback<Uri[]> callback,
                                             FileChooserParams params) {
                return pickFile(callback, params);
            }
        });
        webView.setDownloadListener((url, userAgent, contentDisposition, mimeType, length) ->
                offerDownload(url, fileNameFrom(contentDisposition), mimeType));
        root.removeAllViews();
        root.addView(webView);
        // Сервер сам выдаёт cookie в ответе на этот адрес. setCookie здесь не подходит:
        // она асинхронная, и первая страница могла загрузиться раньше и получить отказ.
        String next = importFile ? "/import/incoming" : "/";
        webView.loadUrl(baseUrl(port) + "unlock?token=" + accessToken + "&next=" + next);
    }

    /** В режиме готовки экран не должен гаснуть. */
    private void updateKeepScreenOn(String url) {
        if (Uri.parse(url).getPath() != null && Uri.parse(url).getPath().endsWith("/cook")) {
            getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
        } else {
            getWindow().clearFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
        }
    }

    private boolean pickFile(ValueCallback<Uri[]> callback, WebChromeClient.FileChooserParams params) {
        if (pendingFilePick != null) {
            pendingFilePick.onReceiveValue(null);
        }
        pendingFilePick = callback;
        Intent intent = params.createIntent();
        // Файлы .json система часто не считает JSON, поэтому для импорта показываем все файлы.
        String[] accept = params.getAcceptTypes();
        boolean image = accept.length > 0 && accept[0].startsWith("image/");
        intent.setType(image ? "image/*" : "*/*");
        Intent request = image ? withCamera(intent) : intent;
        try {
            startActivityForResult(request, REQUEST_PICK_FILE);
            return true;
        } catch (ActivityNotFoundException e) {
            pendingFilePick = null;
            Toast.makeText(this, "Нет приложения для выбора файла", Toast.LENGTH_LONG).show();
            return false;
        }
    }

    /** Добавляет к выбору файла съёмку на камеру. Снимок сохраняется во временный файл. */
    private Intent withCamera(Intent pickFile) {
        pendingCameraPhoto = null;
        Intent chooser = Intent.createChooser(pickFile, "Фото блюда");
        try {
            File dir = new File(getCacheDir(), "camera");
            dir.mkdirs();
            File photo = File.createTempFile("photo", ".jpg", dir);
            Uri uri = FileProvider.getUriForFile(this, FILE_AUTHORITY, photo);
            Intent camera = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
            camera.putExtra(MediaStore.EXTRA_OUTPUT, uri);
            camera.setClipData(ClipData.newRawUri("photo", uri));
            camera.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                    | Intent.FLAG_GRANT_READ_URI_PERMISSION);
            chooser.putExtra(Intent.EXTRA_INITIAL_INTENTS, new Intent[]{camera});
            pendingCameraPhoto = uri;
        } catch (IOException | IllegalArgumentException e) {
            // Без камеры останется обычный выбор файла.
        }
        return chooser;
    }

    private void offerDownload(String url, String name, String mimeType) {
        new AlertDialog.Builder(this)
                .setTitle(name)
                .setItems(new CharSequence[]{"Поделиться", "Сохранить на телефон"}, (dialog, which) -> {
                    if (which == 0) {
                        shareDownload(url, name, mimeType);
                    } else {
                        askWhereToSave(url, name, mimeType);
                    }
                })
                .show();
    }

    /** Запрашиваем у системы, куда сохранить: так не нужны разрешения на доступ к памяти. */
    private void askWhereToSave(String url, String name, String mimeType) {
        pendingDownloadUrl = url;
        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType(mimeType);
        intent.putExtra(Intent.EXTRA_TITLE, name);
        startActivityForResult(intent, REQUEST_SAVE_FILE);
    }

    private void shareDownload(String url, String name, String mimeType) {
        new Thread(() -> {
            try {
                File dir = new File(getCacheDir(), "shared");
                deleteRecursively(dir);
                dir.mkdirs();
                File file = new File(dir, name);
                try (OutputStream out = new FileOutputStream(file)) {
                    download(url, out);
                }
                Uri uri = FileProvider.getUriForFile(this, FILE_AUTHORITY, file);
                Intent send = new Intent(Intent.ACTION_SEND);
                send.setType(mimeType);
                send.putExtra(Intent.EXTRA_STREAM, uri);
                send.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                runOnUiThread(() -> startActivity(Intent.createChooser(send, name)));
            } catch (IOException | IllegalArgumentException e) {
                runOnUiThread(() -> Toast.makeText(this,
                        "Не удалось подготовить файл: " + e.getMessage(), Toast.LENGTH_LONG).show());
            }
        }).start();
    }

    private static void download(String url, OutputStream out) throws IOException {
        HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
        connection.setRequestProperty("Cookie", "access=" + accessToken);
        try (InputStream in = connection.getInputStream()) {
            copy(in, out);
        } finally {
            connection.disconnect();
        }
    }

    private static String fileNameFrom(String contentDisposition) {
        Matcher utf8 = Pattern.compile("filename\\*=UTF-8''([^;]+)").matcher(contentDisposition);
        try {
            if (utf8.find()) {
                return URLDecoder.decode(utf8.group(1), "UTF-8");
            }
        } catch (java.io.UnsupportedEncodingException | IllegalArgumentException ignored) {
            // Ниже вернётся имя попроще.
        }
        Matcher plain = Pattern.compile("filename=\"?([^\";]+)").matcher(contentDisposition);
        return plain.find() ? plain.group(1) : "recept";
    }

    private void saveDownload(String url, Uri target) {
        new Thread(() -> {
            String message = "Файл сохранён";
            try (OutputStream out = getContentResolver().openOutputStream(target)) {
                download(url, out);
            } catch (IOException e) {
                message = "Не удалось сохранить файл: " + e.getMessage();
            }
            String result = message;
            runOnUiThread(() -> Toast.makeText(this, result, Toast.LENGTH_LONG).show());
        }).start();
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == REQUEST_PICK_FILE && pendingFilePick != null) {
            Uri[] result = WebChromeClient.FileChooserParams.parseResult(resultCode, data);
            // Камера не возвращает адрес: снимок лежит в файле, который мы ей дали.
            if (result == null && resultCode == RESULT_OK && pendingCameraPhoto != null) {
                result = new Uri[]{pendingCameraPhoto};
            }
            pendingFilePick.onReceiveValue(result);
            pendingFilePick = null;
            pendingCameraPhoto = null;
        } else if (requestCode == REQUEST_SAVE_FILE && resultCode == RESULT_OK
                && data != null && data.getData() != null && pendingDownloadUrl != null) {
            saveDownload(pendingDownloadUrl, data.getData());
        }
    }

    @Override
    protected void onResume() {
        super.onResume();
        isVisible = true;
    }

    @Override
    protected void onPause() {
        super.onPause();
        isVisible = false;
    }

    private PendingIntent timerIntent() {
        return PendingIntent.getBroadcast(this, 0, new Intent(this, TimerReceiver.class),
                PendingIntent.FLAG_IMMUTABLE | PendingIntent.FLAG_UPDATE_CURRENT);
    }

    /** Методы для страниц книги. Загружаются только наши страницы, см. WebViewClient. */
    private class AppBridge {
        @JavascriptInterface
        public void startTimer(long millis) {
            AlarmManager alarms = getSystemService(AlarmManager.class);
            long at = System.currentTimeMillis() + millis;
            try {
                if (Build.VERSION.SDK_INT < Build.VERSION_CODES.S || alarms.canScheduleExactAlarms()) {
                    PendingIntent show = PendingIntent.getActivity(MainActivity.this, 1,
                            new Intent(MainActivity.this, MainActivity.class),
                            PendingIntent.FLAG_IMMUTABLE);
                    alarms.setAlarmClock(new AlarmManager.AlarmClockInfo(at, show), timerIntent());
                } else {
                    alarms.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, at, timerIntent());
                }
            } catch (SecurityException e) {
                alarms.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, at, timerIntent());
            }
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU
                    && checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS)
                    != PackageManager.PERMISSION_GRANTED) {
                runOnUiThread(() -> requestPermissions(
                        new String[]{Manifest.permission.POST_NOTIFICATIONS}, 0));
            }
        }

        @JavascriptInterface
        public void cancelTimer() {
            getSystemService(AlarmManager.class).cancel(timerIntent());
        }

        @JavascriptInterface
        public void shareText(String text) {
            Intent send = new Intent(Intent.ACTION_SEND);
            send.setType("text/plain");
            send.putExtra(Intent.EXTRA_TEXT, text);
            runOnUiThread(() -> startActivity(Intent.createChooser(send, "Список покупок")));
        }
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }
}
