using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTD.Excel.Model
{
    public class AssetNotificacao
    {
        public AssetNotificacao()
        {
         
        }

        public AssetNotificacao(Asset ativo)
        {
            Ativo = ativo;
        }

        public bool Consultar { get; set; }

        public long Tamanho { get; set; }

        public Asset Ativo { get; set; }
    }
}
