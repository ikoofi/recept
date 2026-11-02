package ru.recept.book;

import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.media.AudioAttributes;
import android.media.RingtoneManager;

/** Срабатывает по будильнику системы: сигнал приходит, даже если приложение свёрнуто или закрыто. */
public class TimerReceiver extends BroadcastReceiver {
    private static final String CHANNEL_ID = "timer";

    @Override
    public void onReceive(Context context, Intent intent) {
        // На экране страница сама подаёт сигнал, второй не нужен.
        if (MainActivity.isVisible) {
            return;
        }
        NotificationManager manager = context.getSystemService(NotificationManager.class);
        NotificationChannel channel = new NotificationChannel(
                CHANNEL_ID, "Таймер готовки", NotificationManager.IMPORTANCE_HIGH);
        channel.setSound(RingtoneManager.getDefaultUri(RingtoneManager.TYPE_ALARM),
                new AudioAttributes.Builder().setUsage(AudioAttributes.USAGE_ALARM).build());
        channel.enableVibration(true);
        channel.setVibrationPattern(new long[]{0, 400, 200, 400, 200, 400});
        manager.createNotificationChannel(channel);

        PendingIntent open = PendingIntent.getActivity(context, 0,
                new Intent(context, MainActivity.class), PendingIntent.FLAG_IMMUTABLE);
        Notification notification = new Notification.Builder(context, CHANNEL_ID)
                .setSmallIcon(android.R.drawable.ic_lock_idle_alarm)
                .setContentTitle("Таймер: время вышло")
                .setContentText("Книга рецептов")
                .setCategory(Notification.CATEGORY_ALARM)
                .setContentIntent(open)
                .setAutoCancel(true)
                .build();
        manager.notify(1, notification);
    }
}
