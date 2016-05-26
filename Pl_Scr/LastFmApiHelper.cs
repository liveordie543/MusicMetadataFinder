using System;
using System.Configuration;
using System.Net.Http;
using System.Web;
using System.Xml;

namespace Pl_Scr
{
    public class LastFmApiHelper
    {
        private static readonly string ApiUrl = ConfigurationManager.AppSettings["ApiUrl"];

        private static readonly string ApiKey = ConfigurationManager.AppSettings["ApiKey"];

        public static string GetReleaseDateBySong(string title, string artist)
        {
            string trackInfoPath = $"?method=track.getInfo&api_key={ApiKey}&artist={HttpUtility.UrlEncode(artist)}&track={HttpUtility.UrlEncode(title)}";
            HttpResponseMessage trackResponse = PerformWebRequest(trackInfoPath, "GET");
            XmlDocument trackXml = new XmlDocument();
            trackXml.LoadXml(trackResponse.Content.ReadAsStringAsync().Result);
            
            string releaseDate = GetReleaseDateFromXml(trackXml);
            return releaseDate == Messages.ReleaseDateNotFound ? GetReleaseDateBySongAlbum(trackXml) : releaseDate;
        }

        private static string GetReleaseDateBySongAlbum(XmlDocument trackXml)
        {
            XmlNode albumNode = trackXml.GetElementsByTagName("album").Item(0);
            if (albumNode != null)
            {
                for (int i = 0; i < albumNode.ChildNodes.Count; i++)
                {
                    XmlNode childNode = albumNode.ChildNodes.Item(i);
                    if (childNode != null && childNode.Name == "mbid")
                    {
                        string albumInfoPath = $"?method=album.getinfo&api_key={ApiKey}&mbid={childNode.InnerText}";
                        HttpResponseMessage albumResponse = PerformWebRequest(albumInfoPath, "GET");
                        XmlDocument albumXml = new XmlDocument();
                        albumXml.LoadXml(albumResponse.Content.ReadAsStringAsync().Result);
                        return GetReleaseDateFromXml(albumXml);
                    }
                }
            }
            return Messages.AlbumInfoNotFound;
        }

        private static string GetReleaseDateFromXml(XmlDocument xml)
        {
            XmlNode releaseDateNode = xml.GetElementsByTagName("releasedate").Item(0);
            string releaseDate = releaseDateNode?.InnerText ?? Messages.ReleaseDateNotFound;
            if (releaseDate == Messages.ReleaseDateNotFound)
            {
                XmlNode tagsNode = xml.GetElementsByTagName("toptags").Item(0) ?? xml.GetElementsByTagName("tags").Item(0);
                if (tagsNode != null)
                {
                    for (int i = 0; i < tagsNode.ChildNodes.Count; i++)
                    {
                        XmlNode tagNode = tagsNode.ChildNodes.Item(i);
                        if (tagNode?.FirstChild != null)
                        {
                            switch (tagNode.FirstChild.InnerText)
                            {
                                case "60s":
                                case "70s":
                                case "80s":
                                case "90s":
                                case "00s":
                                    releaseDate = tagNode.FirstChild.InnerText;
                                    break;
                            }
                            int releaseYear;
                            if (releaseDate == Messages.ReleaseDateNotFound && Int32.TryParse(tagNode.FirstChild.InnerText, out releaseYear))
                            {
                                if (tagNode.FirstChild.InnerText.Length == 4)
                                {
                                    releaseDate = tagNode.FirstChild.InnerText;
                                }
                            }
                            if (releaseDate != Messages.ReleaseDateNotFound)
                            {
                                break;
                            }
                        }
                    }
                }
            }
            return releaseDate;
        }

        private static HttpResponseMessage PerformWebRequest(string operationPath, string httpMethod, HttpContent data = null)
        {
            HttpResponseMessage response = null;
            using (HttpClient client = new HttpClient())
            {
                string fullUrl = ApiUrl + operationPath;
                switch (httpMethod)
                {
                    case "GET":
                        response = client.GetAsync(fullUrl).Result;
                        break;
                    case "POST":
                        response = client.PostAsync(fullUrl, data).Result;
                        break;
                }
            }
            if (response != null && !response.IsSuccessStatusCode)
            {
                throw new HttpRequestException((int)response.StatusCode + response.StatusCode.ToString());
            }
            return response;
        }
    }
}
