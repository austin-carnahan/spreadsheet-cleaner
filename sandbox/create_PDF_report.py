import matplotlib
matplotlib.use('agg')
import matplotlib.pyplot as plt

def create_bar_plot(x, y, title, x_label, y_label, path, description,
                    dated, aspect=1.6, dpi=175, pad=0.1):
    """Create a barplot and save as a .jpg file"""
    plt.style.use('tableau-colorblind10')
    plt.rcParams['font.size'] = 12
    plt.rcParams['axes.labelsize'] = 10
    plt.rcParams['axes.titlesize'] = 10
    plt.rcParams['xtick.labelsize'] = 10
    plt.rcParams['ytick.labelsize'] = 10
    plt.rcParams['figure.titlesize'] = 16
    graph_height, graph_width = plt.figaspect(aspect)
    plt.figure(figsize=(graph_width, graph_height), dpi=dpi)
    plt.bar(x, y, align='center')
    plt.xticks(rotation=15, horizontalalignment='right')
    plt.ylabel(y_label)
    plt.xlabel(x_label)
    plt.title(title)
    plt.tight_layout()
    plt.savefig("{}{}_{}.jpg".format(path, description, dated),
                orientation='landscape', pad_inches=pad)


def create_line_graph(summary_table, title, x_label, y_label,
                      path, description, dated, aspect=1.6, dpi=175, pad=0.1):
    """Create a lineplot and save as a .jpg file"""
    plt.style.use('tableau-colorblind10')
    plt.rc('lines', linewidth=3, linestyle='-', marker='H')
    plt.rcParams['font.size'] = 12
    plt.rcParams['axes.labelsize'] = 10
    plt.rcParams['axes.titlesize'] = 10
    plt.rcParams['xtick.labelsize'] = 10
    plt.rcParams['ytick.labelsize'] = 10
    plt.rcParams['legend.fontsize'] = 6
    plt.rcParams['figure.titlesize'] = 16
    graph_height, graph_width = plt.figaspect(aspect)
    plt.figure(figsize=(graph_width, graph_height), dpi=dpi)
    plt.plot(summary_table)
    plt.title(title)
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    plt.tight_layout()
    plt.ylim(0)
    plt.savefig("{}{}_{}.jpg".format(path, description, dated),
                orientation='landscape', pad_inches=pad)

def create_twoline_graph(summary_table1, summary_table2, summary_label1, summary_label2,
                         title, x_label, y_label, path, description, dated, aspect = 1.6, dpi=175, pad=0.1):
    """Create a lineplot and save as a .jpg file"""
    plt.style.use('tableau-colorblind10')
    plt.rc('lines', linewidth=3, linestyle='-', marker='H')
    plt.rcParams['font.size'] = 12
    plt.rcParams['axes.labelsize'] = 10
    plt.rcParams['axes.titlesize'] = 10
    plt.rcParams['xtick.labelsize'] = 10
    plt.rcParams['ytick.labelsize'] = 10
    plt.rcParams['legend.fontsize'] = 8
    plt.rcParams['figure.titlesize'] = 16
    graph_height, graph_width = plt.figaspect(aspect)
    plt.figure(figsize=(graph_width, graph_height), dpi=dpi)
    plt.plot(summary_table1)
    plt.plot(summary_table2)
    plt.title(title)
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    plt.tight_layout()
    plt.ylim(0)
    plt.legend(loc='best', labels=(summary_label1, summary_label2))
    plt.tight_layout()
    plt.savefig("{}{}_{}.jpg".format(path, description, dated),
                orientation='landscape', pad_inches=pad)


def create_multiline_graph(summary_table, title, x_label, y_label, path,
                           description, dated, offset=0.5, aspect = 1.6, dpi=175, pad=0.1):
    """Create a multi-line lineplot and save as a .jpg file"""
    plt.style.use('tableau-colorblind10')
    plt.rc('lines', linewidth=3, linestyle='-', marker='H')
    plt.rcParams['font.size'] = 12
    plt.rcParams['axes.labelsize'] = 10
    plt.rcParams['axes.titlesize'] = 10
    plt.rcParams['xtick.labelsize'] = 10
    plt.rcParams['ytick.labelsize'] = 10
    plt.rcParams['legend.fontsize'] = 6
    plt.rcParams['figure.titlesize'] = 16
    graph_height, graph_width = plt.figaspect(aspect)
    fig, ax = plt.subplots(figsize=(graph_width, graph_height), dpi=dpi)
    summary_table.unstack().plot(ax=ax)
    plt.title(title)
    plt.ylabel(y_label)
    plt.xlabel(x_label)
    leg = plt.legend(loc='best')
    plt.draw()
    bb = leg.get_bbox_to_anchor().inverse_transformed(ax.transAxes)
    bb.x0 += offset
    bb.x1 += offset
    leg.set_bbox_to_anchor(bb, transform=ax.transAxes)
    plt.tight_layout()
    plt.savefig("{}{}_{}.jpg".format(path, description, dated),
                orientation='landscape', pad_inches=pad)
