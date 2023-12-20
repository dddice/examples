using System;
using System.Net.Http;
using System.Text;

// This is a C# code action for streamer.bot that will roll a d20
// when activated. It can be used as the result of any trigger in
// streamer.bot that allows actions written in C#.
// For more information read the streamer.bot docs at
// https://docs.streamer.bot/guide/csharp
class CPHInline
{
    public bool Execute()
    {
        using (var client = new HttpClient())
        {
            string url = "https://dddice.com/api/1.0/roll";

            // Replace <room-slug-goes-here> with the room slug. You can find this
            // at the end of the url for your room, or the end of the room join link
            // ex. If https://dddice.com/room/a13Tf6e is the room's join link then a13Tf6e
            // is its slug.
            // for how to format the json data below, see our docs at: https://docs.dddice.com/api/#roll-POSTapi-1.0-roll
            string data = "{ \"room\":\"<room-slug-goes-here>\", \"dice\": [{ \"type\": \"d20\", \"theme\": \"witchlight-l9zyi1mi\"}] }";

            // Replace <api-key-goes-here> with your api key
            // you can generate one from https://dddice.com/account/developer
            // click on "Create API Key" and the key will be copied to your clip board
            client.DefaultRequestHeaders.Add("Authorization", "Bearer <api-key-goes-here>");

            var content = new StringContent(data, Encoding.UTF8, "application/json");

            client.PostAsync(url, content).Wait();
        }
        return true;
    }
}