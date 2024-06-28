using System;
using System.Runtime.InteropServices;
using System.IO;
using MicroStationDGN;

namespace MicroStationApp
{
    public class MicroStationApplications
    {
        internal static Application mApp = new MicroStationDGN.Application();

        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        public interface IMicApp
        {
            void OpenFile(string path);
            string GetFileName();
            string GetLevels();
            string SearchTextCoordinate(string oLvlName, string inStringWord);
        }

        [ClassInterface(ClassInterfaceType.None)]
        public class MicApp : IMicApp
        {
            public void OpenFile(string path)
            {
                if (File.Exists(path))
                {
                    if (path.ToLower().Contains(".dgn"))
                    {
                        // Read only
                        mApp.OpenDesignFile(path, true);
                    }
                }
            }

            public string GetFileName()
            {
                try
                {
                    return mApp.ActiveDesignFile.Name;
                }
                catch (Exception)
                {
                    return "Any project was not open.";
                }
            }

            public string GetLevels()
            {
                string o_allLevels = "";
                foreach (Level o_lvl in mApp.ActiveDesignFile.Levels)
                {
                    if (string.IsNullOrEmpty(o_allLevels))
                    {
                        o_allLevels = o_lvl.Name;
                    }
                    else
                    {
                        o_allLevels = o_allLevels + Environment.NewLine + o_lvl.Name;
                    }
                }
                return o_allLevels;
            }

            public string SearchTextCoordinate(string oLvlName, string inStringWord)
            {
                ElementEnumerator oE;
                ElementScanCriteria createSearchCriteria = new ElementScanCriteria();
                Level lvl = null;
                foreach (Level o_lvl in mApp.ActiveDesignFile.Levels)
                {
                    if (o_lvl.Name.ToLower().Contains(oLvlName.ToLower()))
                    {
                        lvl = o_lvl;
                        break;
                    }
                }

                if (lvl != null)
                {
                    createSearchCriteria.ExcludeAllClasses();
                    createSearchCriteria.IncludeLevel(lvl);
                    createSearchCriteria.IncludeType(MsdElementType.msdElementTypeText);
                    oE = mApp.ActiveModelReference.Scan(createSearchCriteria);
                    while (oE.MoveNext())
                    {
                        if (oE.Current.IsTextElement)
                        {
                            if (oE.Current.AsTextElement.Text.Contains(inStringWord))
                            {
                                return "X:" + oE.Current.AsTextElement.get_Origin().X +
                                       " Y:" + oE.Current.AsTextElement.get_Origin().Y +
                                       " Z:" + oE.Current.AsTextElement.get_Origin().Z;
                            }
                        }
                    }
                }
                else
                {
                    return "Element not found.";
                }

                return string.Empty;
            }
        }
    }
}
