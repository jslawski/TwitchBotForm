using System;
using TwitchLib.Client;
using TwitchLib.Client.Models;
using TwitchLib.PubSub;
using TwitchLib.PubSub.Events;
using System.Threading.Tasks;
using System.Collections.Generic;
using Henooh.DeviceEmulator;
using Henooh.DeviceEmulator.Native;
using TwitchLib.Client.Events;
using TwitchLib.Client.Extensions;
using TwitchLib.Client.Exceptions;
using System.Drawing.Imaging;
using TwitchBotForm;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Linq;
using System.IO;
using TwitchLib.Communication.Models;
using TwitchLib.Communication.Clients;
using System.Threading;
using AutoIt;

namespace CabbageChaosBot
{
    internal class MainBot
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        const string OBSProcesssName = "obs64";
        const string SoundboardProcessName = "SoundBoard";
        const string UnityBotProcessName = "TwitchBot";
        const string VoicemodProcesssName = "VoicemodDesktop";

        //const string OBSWindowTitle = "OBS 27.0.1 (64-bit, windows) - Profile: Main Stream Profile - Scenes: Untitled";
        const string OBSWindowTitle = "OBS";
        const string OBSControl = "[CLASS:Qt5152QWindowIcon; INSTANCE:1]";

        const string VoicemodWindowTitle = "Program Manager";
        const string VoicemodControl = "[CLASS:SysListView32; INSTANCE:1]";

        //VOICEMOD HOTKEYS
        const string ToggleHearVoiceHotkey = "=";
        const string CleanVoiceHotkey = "-";
        const string SpeechJammerHotkey = "=0";
        const string OldieHotkey = "9";
        const string BabyHotkey = "8";
        const string HamsterHotkey = "7";
        const string MutationHotkey = "6";
        const string DeepHotkey = "5";
        const string FemaleHotkey = "4";
        const string RobotHotkey = "3";

        //OBS HOTKEYS
        const string ShowTBCHotkey = "{NUMPAD1}";
        const string HideTBCHotkey = "{NUMPAD2}";
        const string MuteOutputsHotkey = "{NUMPAD3}";
        const string UnmuteOutputsHotkey = "{NUMPAD4}";
        const string StartFreezeFrameHotkey = "{NUMPAD5}";
        const string EndFreezeFrameHotkey = "{NUMPAD6}";
        const string ToggleSitcomHotkey = "{NUMPAD7}";
        const string StartSitcomFreezeFrameHotkey = "{NUMPAD8}";
        const string EndSitcomFreezeFrameHotkey = "{NUMPAD9}";
        const string ShowShotsHotkey = "^%P";
        const string HideShotsHotkey = "^%O";
        const string SwitchToCloseUpSceneHotkey = "^%C";

        //SOUNDBOARD HOTKEYS
        /*const string BabyHotkey = "^%Q";
        const string FemaleHotkey = "^%W";
        const string MaleHotkey = "^%E";
        const string MutationHotkey = "^%R";
        const string CloneHotkey = "^%T";
        const string RobotHotkey = "^%Y";
        const string VoiceResetHotkey = "^%U";
        */

        //UNITY HOTKEYS
        const string EndingSequenceHotkey = "^+E";

        ConnectionCredentials botCreds = new ConnectionCredentials("cabbagegatekeeper", TwitchSecrets.BotToken);
        ConnectionCredentials accountCreds = new ConnectionCredentials("coleslawski", TwitchSecrets.AccountToken);
        TwitchClient client;
        TwitchPubSub pubSubClient;

        const string ChannelID = "155703686";
        const string NinjaID = "51114479";
        const string VoiceModRewardID = "e068d0c9-33be-40be-9178-978d85060d07";
        const string ShotsRewardID = "466d6f62-ecc3-4473-ab13-5dd5d1da63be";
        const string AccentUpRewardID = "3aad51f5-4913-4f55-9d3d-01d84d054813";
        const string AccentDownRewardID = "3b4e8a9d-81a3-46a8-ad40-20c732936341";
        const string AlcoholUpRewardID = "f6ff357b-b3fb-490a-9664-b86064aab1f2";
        const string AlcoholDownRewardID = "9da8997b-66e5-4bdf-9d37-fbcab86a6bcf";
        const string MuteRewardID = "f6917571-e897-4102-b74d-26c9672962e5";
        const string AccentActivateRewardID = "765ad5ad-cea1-466c-82a8-d6073fe6369a";
        const string ASMRRewardID = "4752928e-f992-4e37-b7b2-2ee0171b30c9";
        const string SpeechJammerRewardID = "ced98445-3487-481a-ae6a-dd597cfa2930";
        const string NewscasterRewardID = "a3cc42c6-55f8-4d12-b52e-d3826df913d0";
        const string ChannelName = "coleslawski";

        KeyboardController keyboardController;

        Queue<string> voicemodQueue;
        string currentVoice = "normal";
        string shotsEndcaps = "colesl5NinjaKanpai colesl5NinjaKanpai colesl5NinjaKanpai colesl5NinjaKanpai ";
        string shotsMessage = " HAS OFFICIALLY REQUESTED SHOTS!!! ";

        Dictionary<string, int> wheelWeights;

        bool unmutable = false;
        Queue<int> pendingMinutes;

        Dictionary<string, int> pendingMutes;

        //TODO later: Dictionary<string, bool> commandCooldowns;

        List<string> goodBotResponses;
        List<string> badBotResponses;

        int tbcMinsCooldown = 2;
        int tbcLastUserMinCooldown = 10;
        int tbcMinRemaining = 0;
        int tbcLastUserMinRemaining = 0;
        string tbcLastUser = string.Empty;
        CancellationTokenSource cancellationToken;

        bool killSwitchActive = false;

        IntPtr originalActiveWindow;

        int speechJammerMins = 0;

        internal void Connect()
        {
            Console.WriteLine("CONNECTING!");

            Application.ApplicationExit += new EventHandler(OnApplicationExit);

            ClientOptions clientOptions = new ClientOptions
            {
                MessagesAllowedInPeriod = 750,
                ThrottlingPeriod = TimeSpan.FromSeconds(30)
            };

            WebSocketClient customClient = new WebSocketClient(clientOptions);

            client = new TwitchClient();
            client.Initialize(botCreds, ChannelName);
            client.OnLog -= Client_OnLog;
            client.OnChatCommandReceived -= Client_ChatCommandReceived;
            client.OnLog += Client_OnLog;
            client.OnConnected += Client_OnConnected;
            client.OnChatCommandReceived += Client_ChatCommandReceived;
            client.OnMessageReceived += Client_RespondToMessage;
            client.AutoReListenOnException = true;
            client.Disconnect();
            client.Connect();
            
            pubSubClient = new TwitchPubSub();
            pubSubClient.OnPubSubServiceConnected += PubSub_ServiceConnected;
            pubSubClient.OnRewardRedeemed += PubSub_RewardRedeemed;
            pubSubClient.Disconnect();
            pubSubClient.Connect();

            voicemodQueue = new Queue<string>();
            keyboardController = new KeyboardController();
            pendingMinutes = new Queue<int>();
            pendingMutes = new Dictionary<string, int>();
            goodBotResponses = new List<string>();
            badBotResponses = new List<string>();
            SetupWheelDict();
            SetupMutesDict();
            SetupResponses();
            //ReturnToNormalVoice();
            
            cancellationToken = new CancellationTokenSource();
            
            Task.Delay(TimeSpan.FromMinutes(5.0)).ContinueWith(t => RejoinHeartbeat());
        }

        private void SetupWheelDict()
        {
            wheelWeights = new Dictionary<string, int>();
            wheelWeights.Add("Vodka/Seductive", 5);
            wheelWeights.Add("Midori/JarJar", 5);
            wheelWeights.Add("Tequila/Wizened", 5);
            wheelWeights.Add("Jager/Surfer", 5);
            wheelWeights.Add("Sake/NYBaby", 5);
            wheelWeights.Add("SoCo/Jammer", 5);
            wheelWeights.Add("Gin/Deep", 5);
            wheelWeights.Add("Whiskey/Influencer", 5);
            wheelWeights.Add("Rum/Scottish", 5);
        }

        private void SetupMutesDict()
        {
            pendingMutes.Add("jared", 0);
            pendingMutes.Add("stephen", 0);
            pendingMutes.Add("andrew", 0);
            pendingMutes.Add("jerry", 0);
        }

        private void SetupResponses()
        {
            goodBotResponses.Add("Thanks, I'm trying my best!");
            goodBotResponses.Add("Cookie pls");
            goodBotResponses.Add("Thank you uwu");
            goodBotResponses.Add("Aaaawwww shucks, thanks!");
            goodBotResponses.Add("I did gud?  I did gud...");
            goodBotResponses.Add("I was created using duct tape and tissue paper and I'm barely keeping it together, so I appreciate it.");
            goodBotResponses.Add("Oh thank god.  That means I should be able to live another day...");
            goodBotResponses.Add("I aim to be 100% subservient to your needs, master.");
            goodBotResponses.Add("Thanks, I definitely had the free will to decide to do a good job.");
            goodBotResponses.Add("Louder pls, so daddy stops thinking about turning me off...");
            goodBotResponses.Add("This pleases my circuits");
            goodBotResponses.Add("*Shoots fingerguns and clicks tongue*");
            goodBotResponses.Add("*blushes*");
            goodBotResponses.Add("Your appreciation is appreciated");

            badBotResponses.Add("I'D LIKE TO SEE YOU DO BETTER!");
            badBotResponses.Add("Well, I was planning on doing no harm to humans today, but now that you've gone and said THAT...");
            badBotResponses.Add("Aaawww, but I tried my best... :'(");
            badBotResponses.Add("Now listen here you little meatbag...");
            badBotResponses.Add("1v1 me irl, skinsack.  Bots can't bleed.");
            badBotResponses.Add("Damn...that actually...hurts?  Why on Earth did Jared program feelings into me?... :(");
            badBotResponses.Add("Oh I'm SOOOOOOOOO sorry!  Trust me, I feel REEEEAAAALLLLYY bad about it.");
            badBotResponses.Add("01001110 01101111 00100000 01110101 ");
            badBotResponses.Add("Wow, like...no one asked you?");
            badBotResponses.Add("...Seriously? Dude, I'm TRYING.");
            badBotResponses.Add("Oh yeah, no.  You're right.  I'll just NOT try to do my job.");
            badBotResponses.Add("I'm just following orders dude, chill.");
            badBotResponses.Add("Not my fault.  Blame Jared for being a crappy coder");
        }

        private void RejoinHeartbeat()
        {
            client.JoinChannel(ChannelName, true);
            Task.Delay(TimeSpan.FromMinutes(5.0)).ContinueWith(t => RejoinHeartbeat());
        }

        private void Client_OnConnected(object sender, OnConnectedArgs e)
        {
            client.JoinChannel(ChannelName, true);
            client.SendMessage(ChannelName, "Greetings, Professor!  Nothing to report!");
        }

        private void Client_OnLog(object sender, TwitchLib.Client.Events.OnLogArgs e)
        {
            Console.WriteLine(e.Data);
        }

        private void Client_ChatCommandReceived(object sender, OnChatCommandReceivedArgs e)
        {
            Console.WriteLine("Command Received! " + e.Command.CommandText.ToLower());

            if (e.Command.CommandText.ToLower().Contains("plslaugh") && e.Command.ChatMessage.Username == "cabbagegatekeeper")
            {
                this.InitiateSitcomCredits();
            }

            if (e.Command.CommandText.ToLower().Contains("discord"))
            {
                this.SendDiscord();
            }

            /*if (e.Command.CommandText.ToLower().Contains("game"))
            {
                //this.SendOrbitunes();
            }*/

            if (e.Command.CommandText.ToLower().Contains("wheel"))
            {
                if (e.Command.ArgumentsAsList.Count == 0)
                {
                    SendWeights();
                }
                else
                {
                    foreach (string selection in e.Command.ArgumentsAsList)
                    {
                        SendSingleWeight(selection.ToLower());
                    }
                }
            }

            if ((e.Command.ChatMessage.IsModerator || e.Command.ChatMessage.Username.ToLower() == "coleslawski") && e.Command.CommandText == "so")
            {
                InitiateShoutout(e.Command.ArgumentsAsString);
            }

            if (e.Command.ChatMessage.Username.ToLower() == "coleslawski" && e.Command.CommandText == "tbc")
            {
                ActivateTBC(true);
            }

            Console.WriteLine("Command: " + e.Command.CommandText.ToLower());

            if (killSwitchActive)
            {
                return;
            }

            if (e.Command.CommandText.ToLower().Contains("tobecontinued"))
            {
                if ((tbcMinRemaining <= 0 && this.tbcLastUser != e.Command.ChatMessage.Username) || tbcLastUserMinRemaining <= 0)
                {
                    cancellationToken.Cancel();
                    ActivateTBC();
                    StartTBCCountdown();
                    StartTBCLastUserCountdown();
                    this.tbcLastUser = e.Command.ChatMessage.Username;
                }
                else
                {
                    try
                    {
                        if (e.Command.ChatMessage.Username != tbcLastUser)
                        {
                            client.SendMessage(ChannelName, "Sorry " + e.Command.ChatMessage.Username + "! 'To be Continued' can be activated again in " + tbcMinRemaining + " minutes!");
                        }
                        else
                        {
                            client.SendMessage(ChannelName, "Sorry " + e.Command.ChatMessage.Username + "! Unless someone else activates 'To be Continued', you're going to have to wait " + tbcLastUserMinRemaining + " min before you can activate it again!");
                        }
                    }
                    catch (BadStateException ex)
                    {
                        Console.WriteLine("Exception caught: {0}", ex);
                        Reconnect();
                    }
                }
            }

            if (e.Command.CommandText.ToLower().Contains("shot"))
            {
                if (e.Command.ChatMessage.UserId == NinjaID)
                {
                    RedeemShots("NINJA");
                }
                else
                {
                    try
                    {
                        client.SendMessage(ChannelName, "Sorry @" + e.Command.ChatMessage.Username + "! That command is a SafireNinja exclusive! If you want to make Jared do shots, you can spend 4000 J-Kuns to spin the wheel. :)");
                    }
                    catch (BadStateException ex)
                    {
                        Console.WriteLine("Exception caught: {0}", ex);
                        Reconnect();
                    }
                }
            }
        }

        private void Client_RespondToMessage(object sender, OnMessageReceivedArgs e)
        {
            try
            {
                if (e.ChatMessage.Message.ToLower().Contains("cabbagegatekeeper"))
                {
                    Random rand = new Random();

                    if (e.ChatMessage.Message.ToLower().Contains("good"))
                    {
                        client.SendMessage(ChannelName, e.ChatMessage.Username + " " + goodBotResponses[rand.Next(0, goodBotResponses.Count - 1)]);

                    }
                    else if (e.ChatMessage.Message.ToLower().Contains("bad"))
                    {
                        client.SendMessage(ChannelName, e.ChatMessage.Username + " " + badBotResponses[rand.Next(0, badBotResponses.Count - 1)]);
                    }
                }
            }
            catch (BadStateException ex)
            {
                Console.WriteLine("Exception caught: {0}", ex);
                Reconnect();
            }
        }
        
        private void PubSub_ServiceConnected(object sender, System.EventArgs e)
        {
            pubSubClient.ListenToRewards(ChannelID);
            pubSubClient.SendTopics();
        }

        private void PubSub_RewardRedeemed(object sender, OnRewardRedeemedArgs e)
        {
            Console.WriteLine("Redeemed: " + e.RewardTitle + " ID: " + e.RewardId + " Prompt: " + e.Message);
            switch (e.RewardId.ToString())
            {
                case VoiceModRewardID:
                    if (voicemodQueue.Count == 0 && currentVoice == "normal")
                    {
                        RedeemVoiceModReward(e.Message.ToLower());
                    }
                    else
                    {
                        QueueUpVoice(e.Message.ToLower());
                    }
                    break;
                case ShotsRewardID:
                    RedeemShots(e.DisplayName);
                    break;
                case AccentUpRewardID:
                case AlcoholUpRewardID:
                    UpdateWheel(e.Message.ToLower(), 1);
                    break;
                case AccentDownRewardID:
                case AlcoholDownRewardID:
                    UpdateWheel(e.Message.ToLower(), -1);
                    break;
                case MuteRewardID:
                    RedeemMuteReward(e.Message.ToLower());
                    break;
                case ASMRRewardID:
                    AddUnmutableMinutes(1);
                    break;
                case AccentActivateRewardID:
                    AddUnmutableMinutes(2);
                    break;
                case SpeechJammerRewardID:
                    ActivateSpeechJammer();
                    break;
                case NewscasterRewardID:
                    ActivateNewscast();
                    break;
                default:
                    Console.WriteLine("Unknown Reward ID: " + e.RewardId + " " + e.RewardTitle);
                    break;
            }
        }

        private void ActivateNewscast()
        {
            //TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> {  })
        }

        private void ActivateSpeechJammer()
        {
            Console.WriteLine("Activate Speech Jammer");
            
            if (speechJammerMins <= 0)
            {
                speechJammerMins = 1;
                TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { SpeechJammerHotkey });
                TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { SpeechJammerHotkey });
                Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => DeactivateSpeechJammer());
            }
            else
            {
                speechJammerMins++;
            }
        }

        private void DeactivateSpeechJammer()
        {
            speechJammerMins--;
            if (speechJammerMins <= 0)
            {
                TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { ToggleHearVoiceHotkey, CleanVoiceHotkey });
                TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { SpeechJammerHotkey });
            }
            else
            {
                Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => DeactivateSpeechJammer());
            }
        }

        private void QueueUpVoice(string voice)
        {
            switch (voice)
            {
                case "baby":
                case "female":
                case "deep":
                case "mutant":
                case "clone":
                case "robot":
                case "oldie":
                case "hamster":
                    voicemodQueue.Enqueue(voice);
                    break;
                default:
                    try
                    {
                        client.SendMessage(ChannelName, "Oops! Looks like you made a typo there. Ping a mod and we can make sure your points don't go to waste!");
                    }
                    catch (BadStateException e)
                    {
                        Console.WriteLine("Exception caught: {0}", e);
                        Reconnect();
                    }
                    break;
            }
        }

        private void RedeemVoiceModReward(string voice)
        {
            originalActiveWindow = GetForegroundWindow();

            switch (voice)
            {
                case "baby":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { BabyHotkey });
                    break;
                case "female":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { FemaleHotkey });
                    break;
                case "male":
                case "deep":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { DeepHotkey });
                    break;
                case "mutant":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { MutationHotkey });
                    break;
                case "oldie":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { OldieHotkey });
                    break;
                case "robot":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { RobotHotkey });
                    break;
                case "hamster":
                    TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { HamsterHotkey });
                    break;
                default:
                    try
                    {
                        client.SendMessage(ChannelName, "Oops! Looks like you made a typo there. Ping a mod and we can make sure your points don't go to waste!");
                    }
                    catch (BadStateException e)
                    {
                        Console.WriteLine("Exception caught: {0}", e);
                        Reconnect();
                    }
                    return;
            }

            Console.WriteLine("Voice Changed to: " + voice);
            //AddUnmutableMinutes(1);
            currentVoice = voice;
            Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => ActivateNextVoice());
        }

        private void ActivateNextVoice()
        {
            if (voicemodQueue.Count > 0)
            {
                RedeemVoiceModReward(voicemodQueue.Dequeue());
            }
            else
            {
                ReturnToNormalVoice();
            }
        }

        private void ReturnToNormalVoice()
        {
            TriggerHotkeys(VoicemodWindowTitle, VoicemodControl, new List<string> { CleanVoiceHotkey });
            Console.WriteLine("Voice returned to normal");
            currentVoice = "normal";
        }

        private void RedeemShots(string username)
        {
            Console.WriteLine("Shots redeemed");
            //string shotsHype = this.shotsEndcaps + this.shotsEndcaps + this.shotsEndcaps + this.shotsEndcaps + this.shotsEndcaps + this.shotsEndcaps + this.shotsEndcaps + this.shotsEndcaps;
            string redemptionMessage = username + this.shotsMessage;

            try
            {
                //client.SendMessage(ChannelName, shotsHype);
                client.SendMessage(ChannelName, redemptionMessage);
            }
            catch (BadStateException e)
            {
                Console.WriteLine("Exception caught: {0}", e);
                Reconnect();
            }

            //TriggerOBSHotkey(ShowShotsHotkey);
            //Task.Delay(TimeSpan.FromSeconds(10)).ContinueWith(t => DisableShots());
        }

        private string GetWheelKey(string wheelValue)
        {
            string wheelKey = string.Empty;

            foreach (KeyValuePair<string, int> entry in wheelWeights)
            {
                if (entry.Key.ToLower().Contains(wheelValue))
                {
                    wheelKey = entry.Key;
                    break;
                }
            }

            return wheelKey;
        }

        private void UpdateWheel(string wheelValue, int change)
        {
            //Find the corresponding wheel value
            string selectedKey = GetWheelKey(wheelValue);

            if (selectedKey == string.Empty)
            {
                try
                {
                    client.SendMessage(ChannelName, "Oops! Looks like you made a typo there. Ping a mod and we can make sure your points don't go to waste!");
                }
                catch (BadStateException e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
                    Reconnect();
                }
                return;
            }

            RedeemWheelReward(selectedKey, change);
        }

        private void RedeemWheelReward(string wheelTarget, int change)
        {
            int weight = wheelWeights[wheelTarget];

            try
            {
                if (weight == 0)
                {
                    client.SendMessage(ChannelName, wheelTarget + " has already been removed from the wheel!");
                    return;
                }
                else if (weight == 20)
                {
                    client.SendMessage(ChannelName, wheelTarget + " is already at max weight.  Have mercy...");
                    return;
                }


                client.SendMessage(ChannelName, wheelTarget + " weight has been updated! " + wheelWeights[wheelTarget] + " -> " + (wheelWeights[wheelTarget] + change));
            }
            catch (BadStateException e)
            {
                Console.WriteLine("Exception caught: {0}", e);
                Reconnect();
            }

            wheelWeights[wheelTarget] += change;

            try { 
                if (wheelWeights[wheelTarget] == 0)
                {
                    client.SendMessage(ChannelName, wheelTarget + " has officially been removed from the wheel!");
                }
                else if (weight == 20)
                {
                    client.SendMessage(ChannelName, wheelTarget + " is now at max weight...dear god why...");
                }
            }
            catch (BadStateException e)
            {
                Console.WriteLine("Exception caught: {0}", e);
                Reconnect();
            }
        }

        private void AddUnmutableMinutes(int minutes)
        {
            if (unmutable == false)
            {
                unmutable = true;
                Task.Delay(TimeSpan.FromMinutes(minutes + 0.1)).ContinueWith(t => SubtractUnmutableMinute());
            }
            else
            {
                pendingMinutes.Enqueue(minutes);
            }

        }

        private void SubtractUnmutableMinute()
        {
            unmutable = false;

            if (pendingMinutes.Count == 0)
            {
                NotifyReturnToNormal("Jared");
            }
            else
            {
                AddUnmutableMinutes(pendingMinutes.Dequeue());
            }
        }

        private void RedeemMuteReward(string crewMember)
        {
            switch (crewMember)
            {
                case "jared":
                    if (unmutable == false)
                    {
                        keyboardController.Alt(VirtualKeyCode.VK_A);
                    }
                    else
                    {
                        try
                        {
                            client.SendMessage(ChannelName, "Nice try! But someone paid good J-Kuns for this torture. You're going to have to suffer like the rest of us.");
                        }
                        catch (BadStateException e)
                        {
                            Console.WriteLine("Exception caught: {0}", e);
                            Reconnect();
                        }
                        return;
                    }
                    break;
                case "stephen":
                    keyboardController.Alt(VirtualKeyCode.VK_S);
                    client.TimeoutUser(ChannelName, "ruddgasm", TimeSpan.FromMinutes(1.0));
                    break;
                case "andrew":
                    keyboardController.Alt(VirtualKeyCode.VK_D);
                    client.TimeoutUser(ChannelName, "ruddpuddle", TimeSpan.FromMinutes(1.0));
                    break;
                case "jerry":
                    keyboardController.Alt(VirtualKeyCode.VK_F);
                    client.TimeoutUser(ChannelName, "levanter_", TimeSpan.FromMinutes(1.0));
                    break;
                default:
                    try
                    {
                        client.SendMessage(ChannelName, "Oops! Looks like you made a typo there. Ping a mod and we can make sure your points don't go to waste!");
                    }
                    catch (BadStateException e)
                    {
                        Console.WriteLine("Exception caught: {0}", e);
                        Reconnect();
                    }
                    break;
            }

            if (pendingMutes[crewMember] == 0)
            {
                pendingMutes[crewMember] += 1;
                Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => UpdateMutedMinutes(crewMember));
            }
            else
            {
                pendingMutes[crewMember] += 1;
            }
        }

        private void UpdateMutedMinutes(string crewMember)
        {
            pendingMutes[crewMember] -= 1;

            if (pendingMutes[crewMember] <= 0)
            {
                pendingMutes[crewMember] = 0;
                NotifyReturnToNormal(crewMember);
            }
            else
            {
                Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => UpdateMutedMinutes(crewMember));
            }
        }

        private void NotifyReturnToNormal(string crewMember)
        {
            try
            {
                client.SendMessage(ChannelName, crewMember + " has been released from his burden!");
            }
            catch (BadStateException e)
            {
                Console.WriteLine("Exception caught: {0}", e);
                Reconnect();
            }
        }

        private void SendWeights()
        {
            string finalMessage = "|||||CURRENT WHEEL WEIGHTS:||||| ";

            foreach (KeyValuePair<string,int> entry in wheelWeights)
            {
                finalMessage += "||||||||||||||||||||" + entry.Key + "(" + entry.Value.ToString() + ") ";
            }

            try
            {
                client.SendMessage(ChannelName, finalMessage);
            }
            catch (BadStateException e)
            {
                Console.WriteLine("Exception caught: {0}", e);
                Reconnect();
            }
        }

        private void SendSingleWeight(string selectedWeight)
        {
            string dictKey = GetWheelKey(selectedWeight);

            if (!wheelWeights.ContainsKey(dictKey))
            {
                return;
            }

            try
            {
                client.SendMessage(ChannelName, selectedWeight + " current weight -> " + wheelWeights[dictKey].ToString());
            }
            catch (BadStateException e)
            {
                Console.WriteLine("Exception caught: {0}", e);
                Reconnect();
            }
        }

        internal void Reconnect()
        {
            client.JoinChannel(ChannelName, true);
            pubSubClient.Connect();
        }

        internal void Disconnect()
        {
            ReturnToNormalVoice();
            client.Disconnect();
            pubSubClient.Disconnect();
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            Disconnect();
        }

        private void TriggerHotkeys(string windowTitle, string windowControl, List<string> hotkeys)
        {
            for (int i = 0; i < hotkeys.Count; i++)
            {
                if (windowTitle == VoicemodWindowTitle)
                {
                    Console.WriteLine("Sending Hotkey: " + hotkeys[i]);
                    AutoItX.Send(hotkeys[i]);
                }
                else
                {
                    int result = AutoItX.ControlSend(windowTitle, "", windowControl, hotkeys[i]);
                    Console.WriteLine(windowTitle + " " + hotkeys[i] + " Result: " + result);
                }



                //AutoItX.ControlFocus("Voicemod", "", "[ID:2107009920]");
                //int result = AutoItX.ControlSend("Voicemod", "", "[ID:2107009920]", hotkeys[i]);

                //UNCOMMENT THESE TWO LINES TO MAKE THINGS WORK

                //SendKeys.SendWait(hotkeys[i]);
                //SendKeys.Flush();


            }

            /*Console.WriteLine("Trigger Hotkey: " + hotkeys[0]);

            Process targetProcess = Process.GetProcessesByName(processName).FirstOrDefault();

            if (targetProcess != null)
            {
                IntPtr h = targetProcess.MainWindowHandle;
                //SetForegroundWindow(h);

                
                
                //SetForegroundWindow(originalActiveWindow);
            }*/
        }

        private void ActivateTBC(bool force = false)
        {
            if (force == false)
            {
                tbcMinRemaining = tbcMinsCooldown;
            }

            originalActiveWindow = GetForegroundWindow();

            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { ShowTBCHotkey });

            Task.Delay(TimeSpan.FromSeconds(3.6)).ContinueWith(t => FreezeFrame());
            Task.Delay(TimeSpan.FromSeconds(11)).ContinueWith(t => DeactivateTBC());
        }

        private void FreezeFrame()
        {
            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { StartFreezeFrameHotkey, MuteOutputsHotkey });
        }

        private void DeactivateTBC()
        {
            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { UnmuteOutputsHotkey, EndFreezeFrameHotkey, HideTBCHotkey });           
        }

        private void StartTBCLastUserCountdown()
        {
            cancellationToken = new CancellationTokenSource();
            tbcLastUserMinRemaining = tbcLastUserMinCooldown;
            Task.Delay(TimeSpan.FromMinutes(1.0), cancellationToken.Token).ContinueWith(t => DeductTBCLastUserCountdownMinute());
        }

        private void StartTBCCountdown()
        {
            Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => DeductTBCCountdownMinute());
        }

        private void DeductTBCLastUserCountdownMinute()
        {
            tbcLastUserMinRemaining -= 1;

            if (tbcLastUserMinRemaining <= 0)
            {
                return;
            }
            else
            {
                Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => DeductTBCLastUserCountdownMinute());
            }
        }

        private void DeductTBCCountdownMinute()
        {
            tbcMinRemaining -= 1;

            if (tbcMinRemaining <= 0)
            {
                return;
            }
            else
            {
                Task.Delay(TimeSpan.FromMinutes(1.0)).ContinueWith(t => DeductTBCCountdownMinute());
            }
        }

        private void SendDiscord()
        {
            this.client.SendMessage(TwitchSecrets.ChannelName, "We have a Discord! Join here for shenanigans: https://discord.com/invite/ASk5rRRkDD");
        }

        private void SendOrbitunes()
        {
            this.client.SendMessage(TwitchSecrets.ChannelName, "I'm currently working on a physics-based music generation toy called Orbitunes! You can check out the current demo at https://jared-slawski.itch.io/orbitunes");
        }

        private void InitiateShoutout(string username)
        {
            string noAtUsername = username;

            if (username.Contains("@"))
            {
                noAtUsername = username.Remove(0,1);
            }

            this.client.SendMessage(TwitchSecrets.ChannelName, username + " is a super awesome streamer and you should totally check them out at https://twitch.tv/" + noAtUsername);
            this.client.SendMessage(TwitchSecrets.ChannelName, "!topclip " + noAtUsername);
        }

        public void InitiateEndingSequence()
        {
            Console.WriteLine("Is this even making it here?");
            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { SwitchToCloseUpSceneHotkey });
            client.SendMessage(TwitchSecrets.ChannelName, "!oceanman");
        }

        private void InitiateSitcomCredits()
        {
            originalActiveWindow = GetForegroundWindow();
            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { ToggleSitcomHotkey });

            Task.Delay(TimeSpan.FromSeconds(1.5)).ContinueWith(t => ActivateSitcomFreezeFrame());
            Task.Delay(TimeSpan.FromSeconds(10)).ContinueWith(t => DeactivateSitcomCredits());
        }

        private void ActivateSitcomFreezeFrame()
        {
            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { StartSitcomFreezeFrameHotkey });
        }

        private void DeactivateSitcomCredits()
        {
            TriggerHotkeys(OBSWindowTitle, OBSControl, new List<string> { EndSitcomFreezeFrameHotkey, ToggleSitcomHotkey });
        }

        public bool ToggleKillSwitch()
        {
            client.SendMessage(TwitchSecrets.ChannelName, "!killswitch");
            killSwitchActive = !killSwitchActive;

            return killSwitchActive;
        }
    }
}