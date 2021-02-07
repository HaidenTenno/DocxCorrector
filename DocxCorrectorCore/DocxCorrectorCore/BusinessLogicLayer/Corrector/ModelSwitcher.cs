using System;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
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
