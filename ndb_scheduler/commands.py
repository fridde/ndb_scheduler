import click
import ndb_scheduler


@click.command()
@click.argument('file')
def extract_visits_and_fritids(file=None):
    extractor = ndb_scheduler.Extractor(file)
    extractor.extract_visits_as_sql('w')
    extractor.extract_fritids_as_sql('a')
    
    
@click.command()
@click.option('--step', default=None)
@click.option('-s', '--segment', default="2")   
def refine_class_list(step, file, segment):
    ex = ndb_scheduler.Extractor(file)
    if step == "a" or not ex.sheet_exists('step_a'):
        ex.step_a(segment)
    elif step == "b" or not ex.sheet_exists('step_b'):
        ex.step_b()
    elif step == "c" or step is None:
        ex.step_c('a')
    else:
        click.echo("I couldn't find out which step to perform")
