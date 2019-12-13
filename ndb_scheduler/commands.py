import click
import ndb_scheduler


@click.command()
@click.option('-f', '--file', default='eviga_kalendern.xlsx')
@click.option('-e', '--exclude', default=None)
@click.option('-v', '--verbose', is_flag=True)
def extract_visits_and_fritids(file, exclude, verbose):
    ex = ndb_scheduler.Extractor(file, verbose)
    fm = ['w', 'a'] if exclude is None else ['w', 'w']  # file_mode
    if exclude != 'visits':
        ex.extract_visits_as_sql(fm[0])
    if exclude != 'fritids':
        ex.extract_fritids_as_sql(fm[1])
    
    
@click.command()
@click.option('--step', default=None)
@click.option('-f', '--file', default='klasslista.xlsx')
@click.option('-s', '--segment', default="2")
@click.option('-v', '--verbose', is_flag=True)
def refine_class_list(step, file, segment, verbose):
    ex = ndb_scheduler.Extractor(file, verbose)
    if step is None:
        click.echo("Please provide which step to perform")
    if step == "a":
        ex.step_a(segment)
    elif step == "b":
        ex.step_b(segment)
    elif step == "c":
        ex.step_c(segment)
    elif step == "d":
        ex.step_d(segment)
