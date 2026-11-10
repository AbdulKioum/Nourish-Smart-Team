import { useEffect } from 'react';
import { PushNotifications } from '@capacitor/push-notifications';
import { Capacitor } from '@capacitor/core';
import { AuthProvider } from "./context/AuthContext";
import AppRoutes from "./routes/AppRoutes";
import { supabase } from './lib/supabase'; 

export default function App() {
  
  useEffect(() => {
    if (Capacitor.isNativePlatform()) {
      
     
      PushNotifications.requestPermissions().then(result => {
        if (result.receive === 'granted') {
          PushNotifications.register();
        } else {
          console.log('Push notification permission denied');
        }
      });

      
      PushNotifications.addListener('registration', async (token) => {
        console.log('FCM Token: ', token.value);
        
        
        const currentUserId = localStorage.getItem('user_id'); 

        if (currentUserId && token.value) {
          const { error } = await supabase
            .from('users')
            .update({ fcm_token: token.value })
            .eq('id', currentUserId); 

          if (error) {
            console.error('Failed to save FCM token to database:', error);
          } else {
            console.log('FCM token successfully saved to database!');
          }
        }
      });

      // ৩. রেজিস্ট্রেশন এরর হ্যান্ডেলিং
      PushNotifications.addListener('registrationError', (error) => {
        console.error('Error on registration: ', JSON.stringify(error));
      });

      // ৪. নোটিফিকেশনে ক্লিক করলে সেটি হ্যান্ডেল করা
      PushNotifications.addListener('pushNotificationActionPerformed', (notification) => {
        console.log('Push action performed: ', notification);
      });
    }
  }, []);

  return (
    <AuthProvider>
      <AppRoutes />
    </AuthProvider>
  );
}