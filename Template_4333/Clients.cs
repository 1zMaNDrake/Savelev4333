//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Template_4333
{
    using System;
    using System.Collections.Generic;
    
    public partial class Clients
    {
        public int ID { get; set; }
        public string Код_Заказа { get; set; }
        public System.DateTime Дата_создания { get; set; }
        public System.TimeSpan Время_показа { get; set; }
        public int Код_Клиента { get; set; }
        public string Услуги { get; set; }
        public string Статус { get; set; }
        public Nullable<System.DateTime> Дата_закрытия { get; set; }
        public string Время_проката { get; set; }
    }
}
