import Router from 'express-promise-router';

const defaultRouter = Router();

defaultRouter.get('/',function(req, res) {

res.send('It Works !');
})
export default defaultRouter;