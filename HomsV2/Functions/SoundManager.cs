using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using WMPLib;

namespace HomsV2.Functions
{
    class SoundManager
    {
        private WindowsMediaPlayer player;
        public SoundManager() { 
            player = new WindowsMediaPlayer();
        }

        
        public async void PlaySound(string path, bool loop = false, int duration = 10)
        {
            if (player != null)
            {
                player.controls.stop();
                player.close(); // optional: completely release it
            }

            player.URL = path;
            player.settings.setMode("loop", loop);
            player.settings.volume = 100;
            player.controls.play();

            // Wait for specified duration
            await Task.Delay(duration * 1000);

            // Stop after the duration
            player.controls.stop();
        }

        public void StopSound() { 
            player.controls.stop();
        }
    }
}
