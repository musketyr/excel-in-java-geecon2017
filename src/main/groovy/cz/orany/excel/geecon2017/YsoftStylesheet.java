package cz.orany.excel.geecon2017;

import builders.dsl.spreadsheet.builder.api.CanDefineStyle;
import builders.dsl.spreadsheet.builder.api.Stylesheet;

import static builders.dsl.spreadsheet.api.Color.orange;

public class YsoftStylesheet implements Stylesheet {

    @Override
    public void declareStyles(CanDefineStyle stylable) {
        stylable.style("sponsor", st -> st.foreground(orange));
    }

}
