using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

namespace merge
{
    class PPT{
        private uint GetMaxSlideMasterId(SlideMasterIdList slideMaterIdList)
        {
            uint max = 2147483648;
            if (slideMaterIdList != null)
            {
                foreach (SlideMasterId child in slideMaterIdList.Elements<SlideMasterId>())
                {
                    uint id = child.Id;
                    if (id > max)
                    {
                        max = id;
                    }
                }
            }
            return max;
        }
        private uint GetMaxSlideId(SlideIdList slideIdList)
        {
            uint max = 256;
            if (slideIdList != null)
            {
                foreach (SlideId child in slideIdList.Elements<SlideId>())
                {
                    uint id = child.Id;
                    if (id > max)
                    {
                        max = id;
                    }
                }
            }
            return max;
        }

        private void FixSlideLayoutIds(PresentationPart presPart, ref uint uniqueId)
        {
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = (uint)uniqueId;

                }
                slideMasterPart.SlideMaster.Save();
            }
        }

        public void MergePresentationsSlides (string sourceFolderLocation, string sourcePresentation, string destinationFolderLocation, string destPresentation){

            int id = 0;

            // Open Destination presentation
            using (PresentationDocument myDestDeck = PresentationDocument.Open(destinationFolderLocation+destPresentation, true))
            {
                PresentationPart destPresPart = myDestDeck.PresentationPart;

                if (destPresPart.Presentation.SlideIdList == null)
                    destPresPart.Presentation.SlideIdList = new SlideIdList();
                // Open source presentation
                using (PresentationDocument mySourceDeck = PresentationDocument.Open(sourceFolderLocation+sourcePresentation, false))
                {
                    PresentationPart sourcePresPart = mySourceDeck.PresentationPart;

                    uint uniqueId =GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);
                    uint maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);

                    foreach (SlideId slideId in sourcePresPart.Presentation.SlideIdList)
                    {
                        SlidePart sp;
                        SlidePart destSp;
                        SlideMasterPart destMasterPart;
                        string relId;
                        SlideMasterId newSlideMasterId;
                        SlideId newSlideId;

                        id++;
                        sp = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);
                        relId = sourcePresentation.Remove(sourcePresentation.IndexOf('.')) + id;
                        destSp = destPresPart.AddPart<SlidePart>(sp, relId);
                        destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
                        destPresPart.AddPart(destMasterPart);

                        uniqueId++;
                        newSlideMasterId = new SlideMasterId();
                        newSlideMasterId.RelationshipId = destPresPart.GetIdOfPart(destMasterPart);
                        newSlideMasterId.Id = uniqueId;

                        maxSlideId++;
                        newSlideId = new SlideId();
                        newSlideId.RelationshipId = relId;
                        newSlideId.Id = maxSlideId;
                        destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                        destPresPart.Presentation.SlideIdList.Append(newSlideId);


                        FixSlideLayoutIds(destPresPart, ref uniqueId);

                    }
                    destPresPart.Presentation.Save();

                }
            }

            
        }
        
    }

    class Program
    {
        


        static void Main(string[] args)
        {
            PPT program = new PPT();

            //string[] sourcePresentations =Array["ppt/file1.pptx", "ppt/file1.pptx"];
            Console.WriteLine(" File: ppt / file1.pptx");
            Console.WriteLine(File.Exists("ppt/file1.pptx"));

            Console.WriteLine(" File: ppt / file2.pptx");
            Console.WriteLine(File.Exists("ppt/file2.pptx"));



            program.MergePresentationsSlides(@"ppt","\file1.pptx",@"ppt","\file2.pptx");





                

        }

    }
}
