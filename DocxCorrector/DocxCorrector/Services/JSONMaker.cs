﻿using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocxCorrector.Models;

namespace DocxCorrector.Services
{
    public static class JSONMaker
    {
        // Создать JSON строку из списка ошибок mistakes
        public static string MakeMistakesJSON(List<ParagraphResult> results)
        {
            string jsonString = JsonSerializer.Serialize(results);
            return jsonString;
        }
    }
}
