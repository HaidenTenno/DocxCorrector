using System;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel.GOST_7_32;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel.ITMO;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    internal static class ModelSwitcher
    {
        internal static GlobalDocumentModel GetSelectedModel(RulesModel rules)
        {
            return rules switch
            {
                RulesModel.GOST => new DocumentModelGOST_7_32(),
                RulesModel.ITMO => new DocumentModelITMO(),
                _ => throw new NotImplementedException()
            };
        }
    }
}
