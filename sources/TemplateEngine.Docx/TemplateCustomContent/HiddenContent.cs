using System;

namespace TemplateEngine.Docx
{
    public abstract class HiddenContent<TBuilder> : IContentItem
        where TBuilder : HiddenContent<TBuilder>
    {
        protected HiddenContent()
        {
            _instance = (TBuilder) this;
        }

        private readonly TBuilder _instance;

        public TBuilder Hide()
        {
            IsHidden = true;
            return _instance;
        }

        public TBuilder Hide(Func<TBuilder, bool> predicate)
        {
            if (predicate(_instance)) IsHidden = true;

            return _instance;
        }

        public abstract bool Equals(IContentItem other);

        public string Name { get; set; }
        public bool IsHidden { get; set; }
    }
}
