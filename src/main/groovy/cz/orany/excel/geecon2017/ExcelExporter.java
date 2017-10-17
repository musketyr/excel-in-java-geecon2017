package cz.orany.excel.geecon2017;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.builder.api.BorderDefinition;
import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder;
import builders.dsl.spreadsheet.builder.poi.PoiSpreadsheetBuilder;
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria;
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteriaResult;
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static builders.dsl.spreadsheet.api.Color.*;
import static builders.dsl.spreadsheet.api.Keywords.*;

public class ExcelExporter {

    public void basicExport(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        builder.build(w -> w.sheet("Presentations", s -> {
            s.row(r -> {
                r.cell("Name");
                r.cell("Bio");
                r.cell("Twitter");
                r.cell("Title");
                r.cell("Type");
            });

            for (Speaker speaker : speakers) {
                for (Presentation presentation : speaker.getPresentations()) {
                    s.row(r -> {
                        r.cell(speaker.getFirstName() + " " + speaker.getLastName());
                        r.cell(speaker.getShortBio());
                        r.cell(speaker.getTwitter());
                        r.cell(presentation.getTitle());
                        r.cell(presentation.getType());
                    });
                }
            }
        }));
    }

    public void exportWithFilter(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        builder.build(w -> w.sheet("Presentations", s -> {
            s.freeze(1, 1);
            s.filter(auto);
            s.row(r -> {
                r.style(st -> st.font(f -> f.style(bold).size(12)));
                r.cell(c -> c.value("Name").width(auto));
                r.cell(headerCell("Bio"));
                r.cell(headerCell("Twitter"));
                r.cell(headerCell("Title"));
                r.cell(headerCell("Type"));
            });

            for (Speaker speaker : speakers) {
                for (Presentation presentation : speaker.getPresentations()) {
                    s.row(r -> {
                        r.cell(speaker.getFirstName() + " " + speaker.getLastName());
                        r.cell(speaker.getShortBio());
                        r.cell(speaker.getTwitter());
                        r.cell(presentation.getTitle());
                        r.cell(presentation.getType());
                    });
                }
            }
        }));
    }

    public void exportWithRowspan(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        builder.build(w -> w.sheet("Presentations", s -> {
            s.freeze(1, 1);
            s.row(r -> {
                r.style(st -> st.font(f -> f.style(bold).size(12)));
                r.cell(c -> c.value("Name").width(auto));
                r.cell(headerCell("Bio"));
                r.cell(headerCell("Twitter"));
                r.cell(headerCell("Title"));
                r.cell(headerCell("Type"));
            });

            for (Speaker speaker : speakers) {
                s.row(r -> {
                    int numOfPresentations = speaker.getPresentations().size();
                    r.cell(withRowspan(speaker.getFirstName() + " " + speaker.getLastName(), numOfPresentations));
                    r.cell(withRowspan(speaker.getShortBio(), numOfPresentations));
                    r.cell(withRowspan(speaker.getTwitter(), numOfPresentations));
                    r.cell(speaker.getPresentations().get(0).getTitle());
                    r.cell(speaker.getPresentations().get(0).getType());
                });

                if (speaker.getPresentations().size() > 1) {
                    for (int i = 1; i < speaker.getPresentations().size(); i++) {
                        Presentation presentation = speaker.getPresentations().get(i);
                        s.row(r -> {
                            r.cell("D", c -> c.value(presentation.getTitle()));
                            r.cell(c -> c.value(presentation.getType()));
                        });
                    }
                }
            }
        }));
    }

    public void exportWithLinksAndImages(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        builder.build(w -> w.sheet("Presentations", s -> {
            s.freeze(1, 1);
            s.row(r -> {
                r.style(st -> st.font(f -> f.style(bold).size(12)));
                r.cell(headerCell("Name"));
                r.cell(headerCell("Bio"));
                r.cell(headerCell("Twitter"));
                r.cell(headerCell("Title"));
                r.cell(headerCell("Type"));
            });

            for (Speaker speaker : speakers) {
                s.row(r -> {
                    int numOfPresentations = speaker.getPresentations().size();
                    r.cell(withRowspan(speaker.getFirstName() + " " + speaker.getLastName(), numOfPresentations));
                    r.cell(withRowspan(speaker.getShortBio(), numOfPresentations));
                    r.cell(c -> {
                        c.value(speaker.getTwitter());
                        c.rowspan(numOfPresentations);
                        c.style(st -> st.align(center, left));
                        if (speaker.getTwitter() != null) {
                            c.link(to).url("https://twitter.com/" + speaker.getTwitter());
                        }
                    });
                    r.cell(speaker.getPresentations().get(0).getTitle());
                    r.cell(speaker.getPresentations().get(0).getType());
                });

                if (speaker.getPresentations().size() > 1) {
                    for (int i = 1; i < speaker.getPresentations().size(); i++) {
                        Presentation presentation = speaker.getPresentations().get(i);
                        s.row(r -> {
                            r.cell("D", c -> c.value(presentation.getTitle()));
                            r.cell(c -> c.value(presentation.getType()));
                        });
                    }
                }
            }
        }));
    }

    public void exportWithComments(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        builder.build(w -> w.sheet("Presentations", s -> {
            s.freeze(1, 1);
            s.row(r -> {
                r.style(st -> st.font(f -> f.style(bold).size(12)));
                r.cell(headerCell("Name"));
                r.cell(headerCell("Bio"));
                r.cell(headerCell("Twitter"));
                r.cell(headerCell("Title"));
                r.cell(headerCell("Type"));
            });

            for (Speaker speaker : speakers) {
                s.row(r -> {
                    int numOfPresentations = speaker.getPresentations().size();
                    r.cell(withRowspan(speaker.getFirstName() + " " + speaker.getLastName(), numOfPresentations));
                    r.cell(c -> {
                        c.value(speaker.getShortBio());
                        c.rowspan(numOfPresentations);
                        c.style(st1 -> st1.align(center, left));
                        c.comment(speaker.getBio());
                    });
                    r.cell(c -> {
                        c.value(speaker.getTwitter());
                        c.rowspan(numOfPresentations);
                        c.style(st -> st.align(center, left));
                        if (speaker.getTwitter() != null) {
                            c.link(to).url("https://twitter.com/" + speaker.getTwitter());
                        }
                    });
                    Presentation presentation = speaker.getPresentations().get(0);
                    r.cell(presentationDetailsCell(presentation));
                    r.cell(presentation.getType());
                });

                if (speaker.getPresentations().size() > 1) {
                    for (int i = 1; i < speaker.getPresentations().size(); i++) {
                        Presentation presentation = speaker.getPresentations().get(i);
                        s.row(r -> {
                            r.cell("D", presentationDetailsCell(presentation));
                            r.cell(c -> c.value(presentation.getType()));
                        });
                    }
                }
            }
        }));
    }

    private Configurer<CellDefinition> presentationDetailsCell(Presentation presentation) {
        return c -> {
            c.value(presentation.getTitle());
            c.comment(presentation.getDescription());
        };
    }

    private Configurer<CellDefinition> withRowspan(String value, int numOfPresentations) {
        return c -> {
            c.value(value);
            c.rowspan(numOfPresentations);
            c.style(st ->st.align(center, left));
        };
    }

    private Configurer<CellDefinition> headerCell(String title) {
        return c -> {
            c.value(title);
            c.width(auto);
        };
    }

    public void exportPrint(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        builder.build(w -> w.sheet("Presentations", s -> {
            s.page(p -> {
                p.paper(A4);
                p.orientation(portrait);
                p.fit(width);
            });
            for (Speaker speaker : speakers) {
                for (Presentation presentation : speaker.getPresentations()) {
                    s.row(r -> {
                        r.cell(c -> {
                            c.value(speaker.getName());
                            c.width(18).cm();
                            c.height(3).cm();
                            c.style(st -> {
                                st.align(center, center);
                                st.font(f -> {
                                    f.size(15);
                                    f.style(bold);
                                });
                                st.border(top, left, right, b -> {
                                   b.style(thick);
                                   b.color(black);
                                });

                                if (speaker.getShortBio().toLowerCase().contains("y soft")) {
                                    st.foreground(orange);
                                }
                            });
                        });
                    });
                    s.row(r -> {
                        r.cell(c -> {
                            c.value(speaker.getShortBio());
                            c.height(2).cm();
                            c.style(st -> {
                                st.align(center, center);
                                st.font(f -> {
                                    f.size(12);
                                    f.style(italic);
                                });
                                st.border(left, right, b -> {
                                    b.style(thick);
                                    b.color(black);
                                });
                                if (speaker.getShortBio().toLowerCase().contains("y soft")) {
                                    st.foreground(orange);
                                }
                            });
                        });
                    });
                    s.row(r -> {
                        r.cell(c -> {
                            c.value(presentation.getTitle());
                            c.height(5).cm();
                            c.style(st -> {
                                st.align(center, center);
                                st.font(f -> {
                                    f.size(20);
                                    f.style(bold);
                                });
                                st.border(left, right, b -> {
                                    b.style(thick);
                                    b.color(black);
                                });
                                if (speaker.getShortBio().toLowerCase().contains("y soft")) {
                                    st.foreground(orange);
                                }
                            });
                        });
                    });
                    s.row(r -> {
                        r.cell(c -> {
                            c.value(presentation.getDescription());
                            c.height(15).cm();
                            c.style(st -> {
                                st.indent(2);
                                st.wrap(text);
                                st.align(top, justify);
                                st.border(left, right, bottom, b -> {
                                    b.style(thick);
                                    b.color(black);
                                });
                                if (speaker.getShortBio().toLowerCase().contains("y soft")) {
                                    st.foreground(orange);
                                }
                            });
                        });
                    });
                    s.row();
                }
            }
        }));
    }

    public void exportPrintWithStyles(List<Speaker> speakers, File excel) throws FileNotFoundException {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.create(excel);

        String sponsor = "y soft";

        builder.build(w -> {
            if ("y soft".equals(sponsor)) {
                w.apply(YsoftStylesheet.class);
            } else if ("red hat".equals(sponsor)) {
                w.apply(RedHatStylesheet.class);
            }

            w.style("headline", st -> {
                st.font(f -> f.style(bold));
                st.align(center, center);
            });

            w.style("sub", st -> {
                st.font(f -> f.style(italic));
                st.align(center, center);
            });

            w.style("name", st -> {
               st.base("headline");
               st.font(f -> f.size(15));
            });

            w.style("bio", st -> {
                st.base("sub");
                st.font(f -> f.size(12));
            });

            w.style("title", st -> {
                st.base("headline");
                st.font(f -> f.size(20));
            });

            w.style("description", st -> {
                st.indent(2);
                st.wrap(text);
                st.align(top, justify);
            });

            Configurer<BorderDefinition> thickBlackBorder = b -> {
                b.style(thick);
                b.color(black);
            };

            w.style("border-sides", st -> st.border(left, right, thickBlackBorder));

            w.style("border-top", st -> st.border(top, thickBlackBorder));

            w.style("border-bottom", st -> st.border(bottom, thickBlackBorder));

            w.sheet("Presentations", s -> {
                s.page(p -> {
                    p.paper(A4);
                    p.orientation(portrait);
                    p.fit(width);
                });
                for (Speaker speaker : speakers) {
                    for (Presentation presentation : speaker.getPresentations()) {
                        s.row(r -> r.cell(c -> {
                            c.value(speaker.getName());
                            c.width(18).cm();
                            c.height(3).cm();
                            if (speaker.getShortBio().toLowerCase().contains(sponsor)) {
                                c.styles("name", "border-top", "border-sides", "sponsor");
                            } else {
                                c.styles("name", "border-top", "border-sides");
                            }
                        }));
                        s.row(r -> r.cell(c -> {
                            c.value(speaker.getShortBio());
                            c.height(2).cm();
                            if (speaker.getShortBio().toLowerCase().contains(sponsor)) {
                                c.styles("bio", "border-sides", "sponsor");
                            } else {
                                c.styles("bio", "border-sides");
                            }
                        }));
                        s.row(r -> r.cell(c -> {
                            c.value(presentation.getTitle());
                            c.height(5).cm();
                            if (speaker.getShortBio().toLowerCase().contains(sponsor)) {
                                c.styles("title", "border-sides", "sponsor");
                            } else {
                                c.styles("title", "border-sides");
                            }
                        }));
                        s.row(r -> r.cell(c -> {
                            c.value(presentation.getDescription());
                            c.height(15).cm();
                            if (speaker.getShortBio().toLowerCase().contains(sponsor)) {
                                c.styles("description", "border-bottom", "border-sides", "sponsor");
                            } else {
                                c.styles("description", "border-bottom", "border-sides");
                            }
                        }));
                        s.row();
                    }
                }
            });
        });
    }

    public Collection<Speaker> loadSpeakersData(InputStream stream) throws FileNotFoundException {
        SpreadsheetCriteria criteria = PoiSpreadsheetCriteria.FACTORY.forStream(stream);
        SpreadsheetCriteriaResult result = criteria.query(w ->
                w.sheet(s ->
                        s.row(r ->
                                r.cell(c ->
                                        c.style(st -> {
                                            st.font(f ->
                                                    f.size(15)
                                            );
                                            st.foreground(orange);
                                        })))));

        Map<String, Speaker> speakerMap = new LinkedHashMap<>();

        for (Cell cell : result) {
            String name = cell.read(String.class);
            Speaker speaker = speakerMap.get(name);
            if (speaker == null) {
                speaker = new Speaker();
                speaker.setFirstName(name.split("\\s")[0]);
                speaker.setLastName(name.split("\\s")[1]);
                speakerMap.put(name, speaker);
            }

            Cell bioCell = cell.getBellow();
            speaker.setShortBio(bioCell.read(String.class));

            Presentation presentation = new Presentation();
            Cell titleCell = bioCell.getBellow();
            presentation.setTitle(titleCell.read(String.class));

            Cell descriptionCell = titleCell.getBellow();
            presentation.setDescription(descriptionCell.read(String.class));

            speaker.getPresentations().add(presentation);
        }

        return speakerMap.values();
    }


}
